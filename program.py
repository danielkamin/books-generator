import logging
import os
from pathlib import Path
import sys
from configparser import ConfigParser
from datetime import datetime, date
import calendar
import argparse

from db import DatabaseConnection
from document_generator import DocumentGenerator
from services.employee_service import EmployeeService
from utils.paths import find_data_file, resolve_dir


def configure_logging(debug: bool, log_dir: str = "logs"):
    # Ensure log directory exists
    os.makedirs(log_dir, exist_ok=True)

    # Build a timestamped filename, e.g. "2025-08-05-21-00.log"
    timestamp = datetime.now().strftime("%Y-%m-%d-%H-%M")
    logfile = os.path.join(log_dir, f"{timestamp}.log")

    # 1) Root logger catches everything; handlers filter
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    # 2) Console handler
    console_h = logging.StreamHandler(sys.stderr)
    console_h.setLevel(logging.DEBUG if debug else logging.WARNING)
    console_fmt = logging.Formatter("%(asctime)s %(levelname)-8s %(message)s")
    console_h.setFormatter(console_fmt)
    logger.addHandler(console_h)

    # 3) File handler writes ERROR+ to timestamped file
    file_h = logging.FileHandler(logfile, encoding="utf-8")
    file_h.setLevel(logging.ERROR)
    file_fmt = logging.Formatter("%(asctime)s %(name)s %(levelname)-8s %(message)s")
    file_h.setFormatter(file_fmt)
    logger.addHandler(file_h)

    logger.debug(f"Logging initialized. Debug={debug}, logfile={logfile}")


def find_config_file(debug: bool, config_dir: str):
    # 1) first see if user has edited config in ./config/
    base_dir = (
        os.path.dirname(sys.executable)
        if getattr(sys, "frozen", False)
        else os.path.abspath(os.getcwd())
    )
    external = ""
    if debug == True:
        external = os.path.join(base_dir, "config", "config.dev.ini")
    else:
        external = os.path.join(base_dir, "config", "config.ini")

    if os.path.exists(external):
        return external

    # 2) (optional) fall back to a default you did bundle
    #    this only works if you really did bundle a default inside the exe via add-data
    # tempdir = sys._MEIPASS  # where PyInstaller extracted datas
    # return os.path.join(tempdir, 'config', 'config.dev.ini')

    # 3) if neither exists fail early
    raise FileNotFoundError(f"Cannot find config file at {external!r}")


def load_config(debug: bool, config_dir: str):
    path = find_config_file(debug, config_dir)
    cfg = ConfigParser()
    cfg.read(path)
    return cfg


def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS  # type: ignore
    else:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def load_program_config(config: ConfigParser):
    output_directory = config.get("paths", "output_directory", fallback=".")
    supervision_file_name = config.get(
        "paths", "supervision_file_name", fallback="nadzor.csv"
    )
    firearms_file_name = config.get("paths", "firearms_file_name", fallback="bron.csv")
    entities_file_name = config.get(
        "paths", "entities_file_name", fallback="obiekty.csv"
    )

    default_mode = config.get("defaults", "mode", fallback="manual")
    default_interval = config.get("defaults", "interval", fallback="miesieczny")

    return (
        output_directory,
        default_mode,
        default_interval,
        supervision_file_name,
        firearms_file_name,
        entities_file_name,
    )


def choose_option(prompt, options):
    print(prompt)
    for i, option in enumerate(options, 1):
        print(f"{i}. {option}")
    while True:
        try:
            choice = int(input("Select an option by number: "))
            if 1 <= choice <= len(options):
                return options[choice - 1]
            else:
                print("Invalid choice. Try again.")
        except ValueError:
            print("Please enter a valid number.")


def get_current_dates(interval: str):
    today = date.today()
    year = today.year
    month = today.month

    if interval == "miesieczny":
        start = date(year, month, 1)
        end = date(year, month, calendar.monthrange(year, month)[1])
    elif interval == "kwartalny":
        q_start_month = 3 * ((month - 1) // 3) + 1
        q_end_month = q_start_month + 2
        start = date(year, q_start_month, 1)
        end = date(year, q_end_month, calendar.monthrange(year, q_end_month)[1])
    else:
        raise ValueError("Unknown interval")

    return start, end


def manual_date_selection():
    interval = choose_option(
        "Wybierz interwał czasowy:", ["Miesieczny", "Kwartalny"]
    ).lower()

    # Year input with default
    current_year = datetime.now().year
    year_input = input(f"Enter year (default: {current_year}): ").strip()
    year = int(year_input) if year_input else current_year

    q_choice = None
    selected_month = None

    if interval == "miesieczny":
        months = list(calendar.month_name)[1:]  # ['January', ..., 'December']
        selected_month = choose_option("Choose a month:", months)
        month_index = months.index(selected_month) + 1
        start = date(year, month_index, 1)
        end = date(year, month_index, calendar.monthrange(year, month_index)[1])
    else:  # kwartalny
        quarters = ["Kw.1", "Kw.2", "Kw.3", "Kw.4"]
        q_choice = choose_option("Wybierz kwartał:", quarters)
        quarter_index = quarters.index(q_choice) + 1
        q_start_month = 3 * (quarter_index - 1) + 1
        q_end_month = q_start_month + 2
        start = date(year, q_start_month, 1)
        end = date(year, q_end_month, calendar.monthrange(year, q_end_month)[1])

    print(f"📅 Data poczatku: {start}, data konca: {end}")
    return start, end, interval, selected_month, q_choice


def main():
    parser = argparse.ArgumentParser()

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--auto", action="store_true", help="Run in automatic (scheduled) mode"
    )
    parser.add_argument(
        "--data-dir",
        default=None,
        help="Directory with CSV files (defaults to ./data next to the exe)",
    )
    parser.add_argument(
        "--config-dir",
        default=None,
        help="Directory with config files (defaults to ./config next to the exe)",
    )
    parser.add_argument(
        "--debug", action="store_true", help="Show INFO/DEBUG logging to console"
    )
    parser.add_argument(
        "--interval",
        choices=["miesieczny", "kwartalny"],
        default="miesieczny",
        help="Time interval (for --auto mode)",
    )

    args = parser.parse_args()
    DATA_DIR = resolve_dir(args.data_dir, "data")
    CONFIG_DIR = resolve_dir(args.config_dir, "config")
    config = load_config(args.debug, str(CONFIG_DIR))

    (
        output_directory,
        default_mode,
        default_interval,
        supervision_file_name,
        firearms_file_name,
        entities_file_name,
    ) = load_program_config(config)

    configure_logging(debug=args.debug)
    logger = logging.getLogger(__name__)
    logger.info("Starting application")

    # if not os.path.exists(output_directory):
    #     logger.error(f"⚠️ Output directory does not exist: {output_directory}")
    #     raise Exception(
    #         "Output directory for files generation not found, check output_directory config"
    #     )
    # logger.info(f"Using output directory: {output_directory}")

    # init generator values
    start_date = ""
    end_date = ""
    interval = ""
    quarter = ""
    month = ""
    employees = []

    if args.auto:
        start_date, end_date = get_current_dates(args.interval)
        logger.info(f"[AUTO] Generating report from {start_date} to {end_date}")
    else:
        start_date, end_date, interval, month, quarter = manual_date_selection()
        logger.info(f"[MANUAL] Generating report from {start_date} to {end_date}")

    start_date = start_date.strftime("%Y-%m-%d")
    end_date = end_date.strftime("%Y-%m-%d")

    try:
        with DatabaseConnection(config) as db:
            db.connect()
            employee_service = EmployeeService(db)
            employees = employee_service.get_employees_by_period(start_date, end_date)
    except Exception as e:
        print(e)
    finally:
        db.close()

    supervision_file_path = find_data_file(
        supervision_file_name, data_dir=str(DATA_DIR)
    )
    firearms_file_path = find_data_file(firearms_file_name, data_dir=str(DATA_DIR))
    entities_file_path = find_data_file(entities_file_name, data_dir=str(DATA_DIR))

    generator = DocumentGenerator(
        firearms_file_path,
        entities_file_path,
        supervision_file_path,
        month,
        quarter,
        start_date,
        end_date,
        interval,
        employees,
    )

    output_folder_full_path = generator.create_folder_structure()
    if isinstance(output_folder_full_path, Path) == False:
        logger.error(
            "output_folder_full_path from create_folder_structure() is not valid"
        )
        raise Exception(
            "output_folder_full_path from create_folder_structure() is not valid"
        )

    generator.generate_quarterly_reports(output_folder_full_path, start_date, end_date)


if __name__ == "__main__":
    main()

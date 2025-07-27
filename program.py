import os
import sys
from db import DatabaseConnection
from document_generator import DocumentGenerator
from services import EmployeeService
import configparser
from datetime import datetime, date
import calendar
import argparse
from pathlib import Path


def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS  # type: ignore
    else:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def load_program_config(config: configparser.ConfigParser):
    output_directory = config.get("paths", "output_directory", fallback=".")
    default_mode = config.get("defaults", "mode", fallback="manual")
    default_interval = config.get("defaults", "interval", fallback="monthly")

    return output_directory, default_mode, default_interval


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

    if interval == "monthly":
        start = date(year, month, 1)
        end = date(year, month, calendar.monthrange(year, month)[1])
    elif interval == "quarterly":
        q_start_month = 3 * ((month - 1) // 3) + 1
        q_end_month = q_start_month + 2
        start = date(year, q_start_month, 1)
        end = date(year, q_end_month, calendar.monthrange(year, q_end_month)[1])
    else:
        raise ValueError("Unknown interval")

    return start, end


def manual_date_selection():
    interval = choose_option("Choose time interval:", ["Monthly", "Quarterly"]).lower()

    # Year input with default
    current_year = datetime.now().year
    year_input = input(f"Enter year (default: {current_year}): ").strip()
    year = int(year_input) if year_input else current_year

    q_choice = None
    selected_month = None

    if interval == "monthly":
        months = list(calendar.month_name)[1:]  # ['January', ..., 'December']
        selected_month = choose_option("Choose a month:", months)
        month_index = months.index(selected_month) + 1
        start = date(year, month_index, 1)
        end = date(year, month_index, calendar.monthrange(year, month_index)[1])
    else:  # quarterly
        quarters = ["Q1", "Q2", "Q3", "Q4"]
        q_choice = choose_option("Choose a quarter:", quarters)
        quarter_index = quarters.index(q_choice) + 1
        q_start_month = 3 * (quarter_index - 1) + 1
        q_end_month = q_start_month + 2
        start = date(year, q_start_month, 1)
        end = date(year, q_end_month, calendar.monthrange(year, q_end_month)[1])

    print(f"📅 Start date: {start}, End date: {end}")
    return start, end, interval, selected_month, q_choice


def main():
    parser = argparse.ArgumentParser()

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--config",
        default=resource_path("config/config.dev.ini"),
        help="Path to config file",
    )
    parser.add_argument(
        "--auto", action="store_true", help="Run in automatic (scheduled) mode"
    )
    parser.add_argument(
        "--interval",
        choices=["monthly", "quarterly"],
        default="monthly",
        help="Time interval (for --auto mode)",
    )

    args = parser.parse_args()
    config = configparser.ConfigParser()
    config.read(args.config)
    output_directory, default_mode, default_interval = load_program_config(config)

    if not os.path.exists(output_directory):
        print(f"⚠️ Output directory does not exist: {output_directory}")
        raise Exception(
            "Output directory for files generation not found, check output_directory config"
        )
    print(f"Using output directory: {output_directory}")

    start_date = ""
    end_date = ""
    interval = ""
    quarter = ""
    month = ""
    if args.auto:
        start_date, end_date = get_current_dates(args.interval)
        print(f"[AUTO] Generating report from {start_date} to {end_date}")
    else:
        start_date, end_date, interval, month, quarter = manual_date_selection()
        print(f"[MANUAL] Generating report from {start_date} to {end_date}")

    start_date = start_date.strftime("%Y-%m-%d")
    end_date = end_date.strftime("%Y-%m-%d")

    try:
        with DatabaseConnection(config) as db:
            db.connect()
            employee_service = EmployeeService(db)
            employees = employee_service.get_employees_by_period(start_date, end_date)
    except Exception as e:
        print(f"Database error: {e}")
    finally:
        db.close()

    generator = DocumentGenerator(
        month,
        quarter,
        start_date,
        end_date,
        interval,
        DEBUG,
        output_directory,
        resource_path("data/final.csv"),
        resource_path("data/output.csv"),
        resource_path("data/Broń.csv"),
        resource_path("data/supervision.csv"),
    )
    output_folder_full_path = generator.create_folder_structure()
    if isinstance(output_folder_full_path, Path) == False:
        raise Exception(
            "output_folder_full_path from create_folder_structure() is not valid"
        )

    # todo:
    # check if input files exist
    # generate document


if __name__ == "__main__":
    DEBUG = True
    main()

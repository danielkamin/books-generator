from datetime import datetime
from pathlib import Path
import sys
import pandas as pd


class DocumentGenerator:
    """Main class for generating quarterly security service reports."""

    # output.csv to dane pracowników które teraz będą z bazy
    # final.csv to dane o biektach
    # test - Broń.csv to plik z bronią
    # final_csv_path="./data/final.csv",
    #    output_csv_path="./data/output.csv",
    #    firearms_csv_path="./data/Broń.csv",
    def __init__(
        self,
        month,
        quarter,
        start_date="",
        end_date="",
        interval="",
        debug=False,
        output_directory="",
        final_csv_path="",
        output_csv_path="",
        firearms_csv_path="",
        supervision_csv_path="",
    ):
        if (
            output_directory == ""
            or final_csv_path == ""
            or output_csv_path == ""
            or firearms_csv_path == ""
            or supervision_csv_path == ""
            or start_date == ""
            or end_date == ""
            or interval == ""
        ):
            raise Exception(
                "Brakuje którejś ze ściezek do plików źródłowych lub wynikowych"
            )

        self.debug = debug
        self.output_directory = output_directory
        self.start_date = start_date
        self.end_date = end_date
        self.interval = interval
        if month is not None:
            self.month = month
            self.quarter = None
        else:
            self.quarter = quarter
            self.month = None
        self.final_df = self._load_csv(final_csv_path, "final.csv")
        self.employees_df = self._load_csv(output_csv_path, "output.csv")
        self.firearms_df = self._load_csv(firearms_csv_path, "test - Broń.csv")
        self.supervision_data = self._load_csv(supervision_csv_path, "supervision.csv")

        # Convert date columns in employees_df
        if "startdate" in self.employees_df.columns:
            self.employees_df["startdate"] = pd.to_datetime(
                self.employees_df["startdate"], errors="coerce"
            )
        if "enddate" in self.employees_df.columns:
            self.employees_df["enddate"] = pd.to_datetime(
                self.employees_df["enddate"], errors="coerce"
            )

        # Convert date columns in firearms_df
        if "Daty dotyczące przydziału broni palnej" in self.firearms_df.columns:
            self.firearms_df["przydział"] = pd.to_datetime(
                self.firearms_df["Daty dotyczące przydziału broni palnej"],
                errors="coerce",
            )

        # Convert dates in supervision_data
        self.supervision_data["rozpoczęcie"] = pd.to_datetime(
            self.supervision_data["rozpoczęcie"], errors="coerce"
        )
        self.supervision_data["zakończenie"] = pd.to_datetime(
            self.supervision_data["zakończenie"], errors="coerce"
        )

    def _load_csv(self, file_path, name):
        """Helper to load CSV and handle errors."""
        try:
            df = pd.read_csv(file_path, encoding="utf-8")
            self.debug_print(f"Successfully loaded {name} from {file_path}")
            return df
        except FileNotFoundError:
            print(
                f"Error: {name} not found at {file_path}. Please ensure the file exists."
            )
            return pd.DataFrame()  # Return empty DataFrame to avoid further errors
        except Exception as e:
            print(f"Error loading {name} from {file_path}: {e}")
            return pd.DataFrame()

    def debug_print(self, *args, **kwargs):
        """Wrapper for debug prints, using instance debug flag."""
        if self.debug:
            print(*args, **kwargs)

    # def add_page_numbers(self, doc):
    #     """Add page numbers to the document footer."""
    #     section = doc.sections[0]
    #     footer = section.footer
    #     paragraph = (
    #         footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    #     )
    #     paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    #     page_num_run = paragraph.add_run()
    #     page_num_run.font.size = Pt(10)
    #     page_num_run.add_text("Page ")

    #     # Add page number field
    #     page_num_run.add_field("PAGE")  # Use PAGE field for current page number
    #     page_num_run.font.bold = False

    #     page_num_run.add_text(" of ")

    #     # Add total pages field
    #     page_num_run.add_field("NUMPAGES")  # Use NUMPAGES field for total pages
    #     page_num_run.font.bold = False

    #     # Note: Page numbers will only update correctly when the document is opened in MS Word.

    # def get_quarter_dates(self, year, quarter):
    #     """Calculate start and end dates for a given quarter."""
    #     quarters = {
    #         1: ("01-01", "03-31"),
    #         2: ("04-01", "06-30"),
    #         3: ("07-01", "09-30"),
    #         4: ("10-01", "12-31"),
    #     }
    #     start_month_day, end_month_day = quarters[quarter]

    #     quarter_start = pd.to_datetime(f"{year}-{start_month_day}")
    #     quarter_end = pd.to_datetime(f"{year}-{end_month_day}")

    #     return quarter_start, quarter_end

    # def create_document_content(self, doc, year, quarter, department, current_row):
    #     """
    #     Creates a .docx document section based on the provided data for a single contract.
    #     """
    #     try:
    #         quarter_start, quarter_end = self.get_quarter_dates(year, quarter)
    #         poz_ks = int(float(current_row["POZ KS R Umów"]))

    #         # Filter matching rows for Section V (Locations and Form)
    #         matching_rows = self.final_df[
    #             self.final_df["POZ KS R Umów"] == current_row["POZ KS R Umów"]
    #         ].copy()

    #         # I. Księga Realizacji
    #         para = doc.add_paragraph()
    #         para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    #         run = para.add_run("I.  ")
    #         run.bold = True
    #         run.font.size = Pt(14)

    #         run = para.add_run(f"KSIĘGA REALIZACJI UMOWY Nr {poz_ks}")
    #         run.bold = True
    #         run.font.size = Pt(14)

    #         # Add contract party name
    #         para = doc.add_paragraph()
    #         para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    #         para.space_before = Pt(2)
    #         run = para.add_run(
    #             f"    Zawartej z {current_row['Oznaczenie strony lub stron umowy, z którymi przedsiębiorca zawarł umowę']}"
    #         )
    #         run.bold = True
    #         run.font.size = Pt(12)
    #         para.space_after = Pt(12)

    #         # II. Start date
    #         para = doc.add_paragraph()
    #         run = para.add_run(f"II. OD: {quarter_start.strftime('%Y-%m-%d')}")
    #         run.bold = True

    #         # III. End date
    #         para = doc.add_paragraph()
    #         run = para.add_run(f"III. DO: {quarter_end.strftime('%Y-%m-%d')}")
    #         run.bold = True

    #         # IV. Volume (Dział-CK)
    #         para = doc.add_paragraph()
    #         run = para.add_run(
    #             f"IV. VOL. Nr. {current_row['Dział']}-{current_row['CK']}"
    #         )
    #         run.bold = True

    #         # V. Location and form
    #         para = doc.add_paragraph()
    #         run = para.add_run(
    #             "V. Miejsce wykonywania usługi oraz forma jej wykonywania"
    #         )
    #         run.bold = True

    #         table = doc.add_table(rows=len(matching_rows) + 1, cols=8)
    #         table.style = "Table Grid"
    #         col_widths = [5, 8, 22, 22, 18, 9, 9, 7]

    #         for row_obj in table.rows:
    #             for idx, cell in enumerate(row_obj.cells):
    #                 cell.width = Inches(col_widths[idx] / 100 * 7.5)

    #         headers = [
    #             "L.p.",
    #             "Księga",
    #             "Określenie obiektu",
    #             "Adres Obiektu",
    #             "Forma Wykonywanej Usługi",
    #             "Data rozpoczęcia",
    #             "Data zakończenia",
    #             "Uwagi",
    #         ]
    #         for i, header in enumerate(headers):
    #             cell = table.cell(0, i)
    #             cell.text = header
    #             cell.paragraphs[0].runs[0].bold = True

    #         for idx, row_data in matching_rows.iterrows():
    #             row_num = idx - matching_rows.index[0] + 1
    #             data_row = table.rows[row_num]

    #             data_row.cells[0].text = f"{row_num}."
    #             data_row.cells[1].text = str(poz_ks)
    #             data_row.cells[2].text = (
    #                 str(row_data["Określenie obiektu"])
    #                 if pd.notna(row_data["Określenie obiektu"])
    #                 else ""
    #             )
    #             data_row.cells[3].text = (
    #                 str(row_data["Adres Obiektu"])
    #                 if pd.notna(row_data["Adres Obiektu"])
    #                 else ""
    #             )
    #             data_row.cells[4].text = (
    #                 str(row_data["Forma wykonywanej usługi"])
    #                 if pd.notna(row_data["Forma wykonywanej usługi"])
    #                 else ""
    #             )
    #             data_row.cells[5].text = (
    #                 pd.to_datetime(row_data["Data rozpoczęcia usługi"]).strftime(
    #                     "%Y-%m-%d"
    #                 )
    #                 if pd.notna(row_data["Data rozpoczęcia usługi"])
    #                 else ""
    #             )
    #             data_row.cells[6].text = (
    #                 pd.to_datetime(row_data["Data zakończenia usługi"]).strftime(
    #                     "%Y-%m-%d"
    #                 )
    #                 if pd.notna(row_data["Data zakończenia usługi"])
    #                 else ""
    #             )
    #             data_row.cells[7].text = (
    #                 str(row_data["Uwagi"]) if pd.notna(row_data["Uwagi"]) else ""
    #             )

    #         # VI. PRACOWNICY OCHRONY WYKONUJĄCY USŁUGĘ
    #         doc.add_paragraph()
    #         para = doc.add_paragraph()
    #         run = para.add_run("VI. PRACOWNICY OCHRONY WYKONUJĄCY USŁUGĘ")
    #         run.bold = True

    #         current_ck = str(current_row["CK"]).strip()

    #         filtered_employees = self.employees_df[
    #             (self.employees_df["department"].astype(str).str.strip() == current_ck)
    #             & (self.employees_df["startdate"] <= quarter_end)
    #             & (
    #                 (self.employees_df["enddate"].isna())
    #                 | (self.employees_df["enddate"] >= quarter_start)
    #             )
    #         ].copy()

    #         rows_needed = max(
    #             2, len(filtered_employees) + 3
    #         )  # Ensure at least 2 rows for header + 1 empty for appearance
    #         table = doc.add_table(rows=rows_needed, cols=8)
    #         table.style = "Table Grid"

    #         col_widths = [5, 15, 15, 12, 23, 12, 12, 6]
    #         for row_obj in table.rows:
    #             for idx, cell in enumerate(row_obj.cells):
    #                 cell.width = Inches(col_widths[idx] / 100 * 7.5)

    #         headers = [
    #             "L.p.",
    #             "Nazwisko",
    #             "Imię",
    #             "Numer Legitymacji",
    #             "Funkcja w obiekcie",
    #             "Data rozpoczęcia",
    #             "Data zakończenia",
    #             "Uwagi",
    #         ]
    #         for i, header in enumerate(headers):
    #             cell = table.cell(0, i)
    #             cell.text = header
    #             cell.paragraphs[0].runs[0].bold = True

    #         if len(filtered_employees) == 0:
    #             contract_name = current_row[
    #                 "Oznaczenie strony lub stron umowy, z którymi przedsiębiorca zawarł umowę"
    #             ]
    #             self.debug_print(
    #                 f"\nWARNING: No employees found for contract: {contract_name} (CK: {current_ck}, Quarter: Q{quarter} {year})"
    #             )

    #         if self.debug:
    #             self.debug_print(
    #                 f"\nProcessing contract: {current_row['Oznaczenie strony lub stron umowy, z którymi przedsiębiorca zawarł umowę']}"
    #             )
    #             self.debug_print(f"CK value: {current_ck}")
    #             self.debug_print(
    #                 f"Date range: {quarter_start.strftime('%Y-%m-%d')} to {quarter_end.strftime('%Y-%m-%d')}"
    #             )
    #             self.debug_print(f"Found {len(filtered_employees)} matching employees")
    #             if not filtered_employees.empty:
    #                 self.debug_print("First few matching employees:")
    #                 self.debug_print(
    #                     filtered_employees[
    #                         [
    #                             "department",
    #                             "firstname",
    #                             "lastname",
    #                             "startdate",
    #                             "enddate",
    #                         ]
    #                     ].head()
    #                 )

    #         # Format akronim to 5 digits
    #         filtered_employees["akronim"] = filtered_employees["akronim"].apply(
    #             lambda x: str(int(x)).zfill(5) if pd.notna(x) else ""
    #         )

    #         for idx in range(min(len(filtered_employees), rows_needed - 1)):
    #             data_row = table.rows[idx + 1]
    #             emp = filtered_employees.iloc[idx]

    #             start_date_str = (
    #                 emp["startdate"].strftime("%Y-%m-%d")
    #                 if pd.notna(emp["startdate"])
    #                 else ""
    #             )
    #             end_date_str = (
    #                 emp["enddate"].strftime("%Y-%m-%d")
    #                 if pd.notna(emp["enddate"])
    #                 else ""
    #             )

    #             data_row.cells[0].text = f"{idx + 1}."
    #             data_row.cells[1].text = str(emp["lastname"])
    #             data_row.cells[2].text = str(emp["firstname"])
    #             data_row.cells[3].text = str(emp["akronim"])
    #             data_row.cells[4].text = (
    #                 str(emp["function"]) if pd.notna(emp["function"]) else ""
    #             )
    #             data_row.cells[5].text = start_date_str
    #             data_row.cells[6].text = end_date_str
    #             data_row.cells[7].text = ""  # Uwagi column should be empty

    #         for idx in range(
    #             len(filtered_employees), rows_needed - 1
    #         ):  # Fill remaining rows with empty strings
    #             data_row = table.rows[idx + 1]
    #             for cell in data_row.cells:
    #                 cell.text = ""

    #         # VII. PRACOWNICY OCHRONY SPRAWUJĄCY NADZÓR
    #         doc.add_paragraph()
    #         para = doc.add_paragraph()
    #         run = para.add_run(
    #             "VII. PRACOWNICY OCHRONY SPRAWUJĄCY NADZÓR NAD PRACOWNIKAMI OCHRONY WYKONUJĄCYMI USŁUGĘ"
    #         )
    #         run.bold = True

    #         dep_data = self.supervision_data[
    #             self.supervision_data["Dział"] == department
    #         ].copy()
    #         filtered_data = dep_data[
    #             (dep_data["rozpoczęcie"] <= quarter_end)
    #             & (
    #                 (dep_data["zakończenie"].isna())
    #                 | (dep_data["zakończenie"] >= quarter_start)
    #             )
    #         ]

    #         rows_needed_supervision = max(3, len(filtered_data) + 1)
    #         table = doc.add_table(rows=rows_needed_supervision, cols=8)
    #         table.style = "Table Grid"
    #         col_widths = [5, 15, 15, 12, 23, 12, 12, 6]

    #         for row_obj in table.rows:
    #             for idx, cell in enumerate(row_obj.cells):
    #                 cell.width = Inches(col_widths[idx] / 100 * 7.5)

    #         headers = [
    #             "L.p.",
    #             "Nazwisko",
    #             "Imię",
    #             "Numer Legitymacji",
    #             "Funkcja w obiekcie",
    #             "Daty rozpoczęcie",
    #             "Daty zakończenie",
    #             "Uwagi",
    #         ]
    #         for i, header in enumerate(headers):
    #             cell = table.cell(0, i)
    #             cell.text = header
    #             cell.paragraphs[0].runs[0].bold = True

    #         for idx in range(min(len(filtered_data), rows_needed_supervision - 1)):
    #             data_row = table.rows[idx + 1]
    #             row_s = filtered_data.iloc[idx]

    #             start_date_s = ""
    #             if pd.notna(row_s["rozpoczęcie"]):
    #                 # Use the later of the quarter start or the actual start date
    #                 start_date_s = max(row_s["rozpoczęcie"], quarter_start).strftime(
    #                     "%Y-%m-%d"
    #                 )

    #             end_date_s = ""
    #             if pd.notna(row_s["zakończenie"]):
    #                 end_date_s = row_s["zakończenie"].strftime("%Y-%m-%d")
    #             else:
    #                 # If no end date, use quarter end date
    #                 end_date_s = quarter_end.strftime("%Y-%m-%d")

    #             data_row.cells[0].text = f"{idx + 1}."
    #             data_row.cells[1].text = str(row_s["Nazwisko"])
    #             data_row.cells[2].text = str(row_s["Imię"])
    #             data_row.cells[3].text = str(row_s["Nr legitymacji"]).strip('"')
    #             data_row.cells[4].text = str(row_s["Funkcja w obiekcie"])
    #             data_row.cells[5].text = start_date_s
    #             data_row.cells[6].text = end_date_s
    #             data_row.cells[7].text = (
    #                 str(row_s["Uwagi"]) if pd.notna(row_s["Uwagi"]) else ""
    #             )

    #         for idx in range(
    #             len(filtered_data), rows_needed_supervision - 1
    #         ):  # Fill remaining with empty strings
    #             data_row = table.rows[idx + 1]
    #             for cell in data_row.cells:
    #                 cell.text = ""

    #         # VIII. ILOŚĆ I RODZAJ BRONI PALNEJ
    #         doc.add_paragraph()
    #         para = doc.add_paragraph()
    #         run = para.add_run(
    #             "VIII. ILOŚĆ I RODZAJ BRONI PALNEJ PRZYDZIELONEJ PRACOWNIKOM OCHRONY DO WYKONANIA USŁUGI"
    #         )
    #         run.bold = True

    #         table = doc.add_table(rows=2, cols=9)
    #         table.style = "Table Grid"
    #         col_widths = [5, 12, 10, 8, 8, 27, 12, 12, 6]

    #         for row_obj in table.rows:
    #             for idx, cell in enumerate(row_obj.cells):
    #                 cell.width = Inches(col_widths[idx] / 100 * 7.5)

    #         headers = [
    #             "L.p.",
    #             "Rodzaj broni palnej",
    #             "Marka broni",
    #             "Kaliber",
    #             "Ilość",
    #             "Obiekt, do którego przydzielono pracownikom broń palną",
    #             "Przydział broni palnej",
    #             "Cofnięcie przydziału broni palnej",
    #             "Uwagi",
    #         ]

    #         for i, header in enumerate(headers):
    #             cell = table.cell(0, i)
    #             cell.text = header
    #             cell.paragraphs[0].runs[0].bold = True

    #         data_row_firearms = table.rows[1]
    #         data_cells_firearms = ["1", "Nie przyznano", "", "", "", "", "", "", ""]

    #         if department == "MON":
    #             mon_firearms = self.firearms_df[
    #                 (
    #                     self.firearms_df[
    #                         "Obiekt, do którego przydzielono pracownikom broń palną"
    #                     ]
    #                     == "MON"
    #                 )
    #                 & (self.firearms_df["przydział"].notna())
    #             ].copy()

    #             if not mon_firearms.empty:
    #                 row_f = mon_firearms.iloc[0]

    #                 assignment_date_f = ""
    #                 if pd.notna(row_f["przydział"]):
    #                     # If assignment date is before quarter start, use quarter start
    #                     assignment_date_f = max(
    #                         row_f["przydział"], quarter_start
    #                     ).strftime("%Y-%m-%d")

    #                 # For end date, use quarter end date
    #                 end_date_f = quarter_end.strftime("%Y-%m-%d")

    #                 data_cells_firearms = [
    #                     "1",
    #                     (
    #                         str(row_f["Rodzaj broni palnej"])
    #                         if pd.notna(row_f["Rodzaj broni palnej"])
    #                         else ""
    #                     ),
    #                     (
    #                         str(row_f["Marka broni"])
    #                         if pd.notna(row_f["Marka broni"])
    #                         else ""
    #                     ),
    #                     str(row_f["Kaliber"]) if pd.notna(row_f["Kaliber"]) else "",
    #                     str(row_f["Ilość"]) if pd.notna(row_f["Ilość"]) else "",
    #                     (
    #                         str(
    #                             row_f[
    #                                 "Obiekt, do którego przydzielono pracownikom broń palną"
    #                             ]
    #                         )
    #                         if pd.notna(
    #                             row_f[
    #                                 "Obiekt, do którego przydzielono pracownikom broń palną"
    #                             ]
    #                         )
    #                         else ""
    #                     ),
    #                     assignment_date_f,
    #                     end_date_f,
    #                     str(row_f["Uwagi"]) if pd.notna(row_f["Uwagi"]) else "",
    #                 ]

    #         for col, value in enumerate(data_cells_firearms):
    #             data_row_firearms.cells[col].text = value

    #         # IX. ILOŚĆ I RODZAJ ŚRODKÓW PRZYMUSU BEZPOŚREDNIEGO
    #         doc.add_paragraph()
    #         para = doc.add_paragraph()
    #         run = para.add_run(
    #             "IX. ILOŚĆ I RODZAJ ŚRODKÓW PRZYMUSU BEZPOŚREDNIEGO PRZYDZIELONYCH PRACOWNIKOM OCHRONY DO WYKONANIA USŁUGI"
    #         )
    #         run.bold = True

    #         table = doc.add_table(rows=2, cols=6)
    #         table.style = "Table Grid"
    #         col_widths = [5, 35, 10, 25, 15, 10]

    #         for row_obj in table.rows:
    #             for idx, cell in enumerate(row_obj.cells):
    #                 cell.width = Inches(col_widths[idx] / 100 * 7.5)

    #         headers = [
    #             "L.p.",
    #             "Rodzaj środka przymusu bezpośredniego",
    #             "Ilość",
    #             "Obiekt do którego przydzielono pracownikom ś.p.b.",
    #             "Daty przydziału",
    #             "Uwagi",
    #         ]

    #         for i, header in enumerate(headers):
    #             cell = table.cell(0, i)
    #             cell.text = header
    #             cell.paragraphs[0].runs[0].bold = True

    #         data_row_coercion = table.rows[1]
    #         data_cells_coercion = ["1", "Nie przyznano", "", "", "", ""]

    #         for col, value in enumerate(data_cells_coercion):
    #             data_row_coercion.cells[col].text = value

    #     except Exception as e:
    #         print(
    #             f"Error creating document content for POZ KS R Umów {current_row['POZ KS R Umów']}: {e}"
    #         )
    #         raise

    # todo: change base_output_path to full output path for a file
    # def generate_quarterly_reports(self, base_output_path):
    #     """Generates quarterly documents with all records."""
    #     try:
    #         current_year = datetime.now().year

    #         for year in range(2023, current_year + 1):
    #             for quarter in range(1, 5):
    #                 quarter_folder = f"Q{quarter}"
    #                 quarter_path = base_output_path / str(year) / quarter_folder
    #                 quarter_path.mkdir(parents=True, exist_ok=True)

    #                 doc = Document()

    #                 # Set document margins
    #                 for section in doc.sections:
    #                     section.left_margin = Inches(0.5)
    #                     section.right_margin = Inches(0.5)
    #                     section.top_margin = Inches(0.5)
    #                     section.bottom_margin = Inches(1)

    #                 # Add page numbers
    #                 self.add_page_numbers(doc)  # Correctly call as a method

    #                 quarter_start, quarter_end = self.get_quarter_dates(year, quarter)

    #                 # Process all records for the quarter
    #                 # Ensure 'POZ KS R Umów' and 'Dział' are treated consistently
    #                 filtered_contracts = self.final_df[
    #                     self.final_df["POZ KS R Umów"].notna()
    #                     & (self.final_df["POZ KS R Umów"].astype(str) != "")
    #                 ].copy()

    #                 for dept in ["MON", "OFS"]:
    #                     dept_rows = filtered_contracts[
    #                         filtered_contracts["Dział"] == dept
    #                     ].copy()
    #                     self.debug_print(
    #                         f"Processing {dept} department for Q{quarter} {year} - Found {len(dept_rows)} records relevant to department."
    #                     )

    #                     for idx, row in dept_rows.iterrows():
    #                         contract_start_date = pd.to_datetime(
    #                             row["Data rozpoczęcia usługi"], errors="coerce"
    #                         )
    #                         contract_end_date = pd.to_datetime(
    #                             row["Data zakończenia usługi"], errors="coerce"
    #                         )

    #                         # Filter contracts based on their start and end dates relative to the quarter
    #                         # A contract is relevant if it started by the end of the quarter AND (it has no end date OR it ends after the quarter started)
    #                         is_active_in_quarter = (
    #                             pd.notna(contract_start_date)
    #                             and contract_start_date <= quarter_end
    #                         ) and (
    #                             pd.isna(contract_end_date)
    #                             or contract_end_date >= quarter_start
    #                         )

    #                         if is_active_in_quarter:
    #                             try:
    #                                 if (
    #                                     doc.paragraphs and len(doc.paragraphs) > 1
    #                                 ):  # Add page break only if there's existing content
    #                                     doc.add_page_break()

    #                                 self.create_document_content(
    #                                     doc, year, quarter, dept, row
    #                                 )

    #                             except Exception as e:
    #                                 print(
    #                                     f"Error processing document for {dept} - POZ KS R Umów {row['POZ KS R Umów']}: {e}"
    #                                 )

    #                 file_name = f"Quarterly_Report_Q{quarter}_{year}.docx"
    #                 file_path = quarter_path / file_name
    #                 doc.save(file_path)
    #                 print(f"Created quarterly document: {file_path}")

    #     except Exception as e:
    #         print(f"Error in document generation process: {e}")

    def create_folder_structure(self):
        output_path = Path(self.output_directory)

        if output_path.exists() == False:
            output_path.mkdir(parents=True, exist_ok=True)
            return self.create_folder_structure()

        start_dt = datetime.strptime(self.start_date, "%Y-%m-%d")
        end_dt = datetime.strptime(self.end_date, "%Y-%m-%d")
        if start_dt.year != end_dt.year:
            sys.exit(1)

        year_path = output_path / str(start_dt.year)

        if not year_path.exists():
            year_path.mkdir(parents=True, exist_ok=True)
            self.debug_print(f"Folder '{year_path}' created.")
        else:
            self.debug_print(f"Folder '{year_path}' already exists.")

        if self.interval == "monthly" and self.month != None:
            month_folder = year_path / f"{start_dt.month}"
            if not month_folder.exists():
                month_folder.mkdir()
                self.debug_print(f"Folder '{month_folder}' created.")
            else:
                self.debug_print(f"Folder '{month_folder}' already exists.")

            return month_folder

        elif self.interval == "quarterly" and self.quarter != None:
            quarter_folder = year_path / f"{self.quarter}"
            if not quarter_folder.exists():
                quarter_folder.mkdir()
                self.debug_print(f"Folder '{quarter_folder}' created.")
            else:
                self.debug_print(f"Folder '{quarter_folder}' already exists.")

            return quarter_folder

        else:
            self.debug_print(
                "No valid interval (monthly/quarterly) or missing variable."
            )
            sys.exit(1)

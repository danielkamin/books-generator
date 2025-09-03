from dataclasses import dataclass
from typing import Optional, List


@dataclass
class Employee:
    """Employee data model"""

    id: int
    name: str
    kod: str
    _project_code: Optional[str] = None  # internal attribute
    _position: Optional[str] = None
    _ck: Optional[str] = None
    _department_position_map = {
        "DZTO_T1": "Technik PS",
        "DZTO_T2": "technik serwisant",
        "HQ_MK": "Zarząd",
        "HQ_SEK": "Sekretariat",
        "MONMN_JK": "Manager Jakub Kamiński",
        "MONO_ZI": "Monitoring załoga interwencyjna",
        "MONO_OP": "Monitoring operator",
        "MONO_AS": "Monitoring asystent",
        "OFSMN_PK": "Manager Piotr Klimiuk",
        "UPCMN_PCH": "Manager Piotr Chmura",
        "OFSO_": "Pracownik Ochrony",
        "brak": "brak",
    }
    _release_date = ""

    def __str__(self):
        return f"{self.name}, Kod: {self.kod}, Projekt: {self.project_code or '-'}, Funkcja: {self.position}"

    @property
    def project_code(self):
        return self._project_code

    @property
    def position(self):
        return self._position

    @property
    def ck(self):
        return self._ck

    @property
    def first_name(self):
        return self.name.split(" ")[0]

    @property
    def last_name(self):
        return self.name.split(" ")[1]

    @property
    def release_date(self):
        return self._release_date

    @project_code.setter
    def project_code(self, value: Optional[str]):
        if value and value != "brak" and value.startswith("P_"):
            self._project_code = value[2:]  # usuń P_
        else:
            self._project_code = value

    @position.setter
    def position(self, proj_value: str):
        if proj_value.startswith("OFSO_"):
            self._position = self._department_position_map.get("OFSO_")
        else:
            self._position = self._department_position_map.get(proj_value)

    @ck.setter
    def ck(self, proj_value: str):
        if proj_value.startswith("OFSO_"):
            self._ck = proj_value
        elif proj_value.startswith("MON"):
            self._ck = "MON"
        else:
            self._ck = None

    @release_date.setter
    def release_date(self, release_date: str):
        self._release_date = release_date


class EmployeeFactory:
    """Factory class to create Employee objects from database results"""

    @staticmethod
    def create_from_db_result(db_row) -> Employee:
        emp = Employee(id=db_row[0], name=db_row[1], kod=db_row[2])
        if len(db_row) > 3:
            emp.project_code = db_row[3]
            emp.position = emp.project_code
            emp.ck = emp.project_code
        if len(db_row) > 5:
            if db_row[5] == 1:
                emp.release_date = str(db_row[4].date())
            else:
                emp.release_date = ""
        return emp

    @staticmethod
    def create_multiple_from_db_results(db_rows) -> List[Employee]:
        return [EmployeeFactory.create_from_db_result(row) for row in db_rows]

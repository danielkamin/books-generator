from dataclasses import dataclass
from datetime import date
from typing import Optional, List


@dataclass
class Employee:
    """Employee data model"""
    id: int
    name: str
    kod: str
    project_code: Optional[str] = None
    
    def __str__(self):
        return f"{self.name}, Kod: {self.kod}, Projekt: {self.project_code or '-'}"
    
class EmployeeFactory:
    """Factory class to create Employee objects from database results"""
    
    @staticmethod
    def create_from_db_result(db_row) -> Employee:
        print(db_row)
        """Create Employee from database row tuple"""
        return Employee(
            id=db_row[0],
            name=db_row[1],
            kod=db_row[2],
            project_code=db_row[3] if len(db_row) > 3 else None
        )
    
    @staticmethod
    def create_multiple_from_db_results(db_rows) -> List[Employee]:
        """Create list of Employees from database results"""
        return [EmployeeFactory.create_from_db_result(row) for row in db_rows]
from typing import List
from repositories import EmployeeRepository
from models import Employee, EmployeeFactory


class EmployeeService:
    """Service layer for employee business logic"""

    def __init__(self, db_connection):
        self.repository = EmployeeRepository(db_connection)

    def get_employees_by_period(self, start_date: str, end_date: str) -> List[Employee]:
        """
        Get employees with project codes for a given period

        Args:
            start_date: Start date in format 'YYYY-MM-DD'
            end_date: End date in format 'YYYY-MM-DD'

        Returns:
            List of Employee objects
        """
        try:
            db_results = self.repository.get_employees_with_project_codes(
                start_date, end_date
            )
            return EmployeeFactory.create_multiple_from_db_results(db_results)
        except Exception as e:
            print(f"Error in employee service: {e}")
            return []

    def get_all_active_employees(self) -> List[Employee]:
        """
        Get all currently active employees
        """
        try:
            db_results = self.repository.get_active_employees()
            employees = []
            for row in db_results:
                employee = Employee(id=row[0], name=row[1], kod=row[2])
                employees.append(employee)
            return employees
        except Exception as e:
            print(f"Error getting active employees: {e}")
            return []


"""
Repository layer for data access operations.

This module contains repository classes that handle database queries
and data retrieval operations.
"""

from .employee_repository import EmployeeRepository

# Export main classes for easy importing
__all__ = [
    'EmployeeRepository',
]

# Version info
__version__ = '1.0.0'
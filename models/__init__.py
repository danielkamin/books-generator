"""
Data models and entities.

This module contains data classes and factory classes for creating
objects from database results.
"""

from .employee import (
    Employee,
    EmployeeFactory,
)

# Export main classes for easy importing
__all__ = [
    # Data models
    'Employee',
    
    # Factory classes
    'EmployeeFactory',
]

# Version info
__version__ = '1.0.0'
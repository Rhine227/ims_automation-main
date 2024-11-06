"""Data models for IMS automation.

This module defines the data structures used to represent IMS templates and their content.

TODO: Add validation for input values
TODO: Add support for template versioning
TODO: Add methods for data comparison between templates
"""

from dataclasses import dataclass, field
from typing import Dict, List


@dataclass
class Task:
    """Represents a maintenance task.
    
    Attributes:
        name: Task identifier
        description: Task details
        inputs: Dictionary mapping cell coordinates to input values
    """
    name: str
    description: str = ""
    inputs: Dict[str, str] = field(default_factory=dict)

    def to_dict(self) -> dict:
        """Convert task to dictionary representation."""
        return {
            "name": self.name,
            "description": self.description,
            "inputs": self.inputs
        }

@dataclass
class Category:
    """
    Represents a group of related tasks.
    
    Attributes:
        name: Category identifier
        tasks: List of tasks in this category
    """
    name: str
    tasks: List[Task] = field(default_factory=list)

@dataclass
class SheetData:
    """
    Represents processed worksheet data.
    
    Attributes:
        name: Worksheet name
        categories: List of categories in the sheet
    """
    name: str
    categories: List[Category] = field(default_factory=list)

    def to_dict(self):
        """Convert the SheetData instance to a dictionary."""
        return {
            "name": self.name,
            "categories": [category.to_dict() for category in self.categories]
        }

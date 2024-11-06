"""Excel file format converter module.

This module handles the conversion of Excel files from .xls to .xlsx format.
It provides functionality for both single file and directory-wide conversions.

TODO: Add progress callback for batch operations
TODO: Add validation of input/output paths
TODO: Add support for conversion options (e.g., preserve formatting)
"""

from pathlib import Path
import logging
from typing import Union
import pandas as pd

logger = logging.getLogger(__name__)

class ExcelConverter:
    """Handles conversion of Excel files from .xls to .xlsx format."""
    
    @staticmethod
    def convert_file(input_path: Union[str, Path], output_path: Union[str, Path]) -> None:
        """Convert a single .xls file to .xlsx format.
        
        Args:
            input_path: Path to source .xls file
            output_path: Path where .xlsx file should be saved
            
        Raises:
            Exception: If conversion fails
        """
        try:
            df = pd.read_excel(input_path, sheet_name=None)
            with pd.ExcelWriter(output_path) as writer:
                for sheet_name, data in df.items():
                    data.to_excel(writer, sheet_name=sheet_name, index=False)
            logger.info(f"Converted {input_path} to {output_path}")
        except Exception as e:
            logger.error(f"Failed to convert {input_path}: {str(e)}")
            raise

    def process_directory(self, directory: Union[str, Path]) -> None:
        """Process all .xls files in directory.
        
        Args:
            directory: Path to directory containing .xls files
        """
        directory = Path(directory)
        for xls_path in directory.rglob("*.xls"):
            xlsx_path = xls_path.with_suffix('.xlsx')
            self.convert_file(xls_path, xlsx_path)
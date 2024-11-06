import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch
from src.ui.task_completion import TaskCompletionUI

class TestTaskCompletionUI(unittest.TestCase):

    def setUp(self):
        """Set up test fixtures."""
        self.task_completion_ui = TaskCompletionUI()
        self.task_completion_ui.data = [
            {
                'name': 'Sheet1',
                'categories': [
                    {
                        'name': 'Category1',
                        'tasks': [
                            {'name': 'Task1', 'inputs': {'A1': 'input1', 'B1': 'input2'}},
                            {'name': 'Comments:', 'inputs': {'A2': 'input3', 'B2': 'input4'}}
                        ]
                    }
                ]
            }
        ]
        self.task_completion_ui.current_sheet = 0
        self.task_completion_ui.current_category = 0
        self.task_completion_ui.current_task = 0

    @patch('src.ui.task_completion.openpyxl')
    def test_preserve_header(self, mock_openpyxl):
        """Test preserving header rows and columns."""
        source_ws = MagicMock()
        target_ws = MagicMock()
        self.task_completion_ui._preserve_header(source_ws, target_ws)
        # Add assertions to verify the behavior

    @patch('src.ui.task_completion.openpyxl')
    def test_clear_previous_month_data(self, mock_openpyxl):
        """Test clearing previous month data."""
        workbook = MagicMock()
        self.task_completion_ui._clear_previous_month_data(workbook)
        # Add assertions to verify the behavior

    def test_get_total_tasks(self):
        """Test getting total tasks."""
        total_tasks = self.task_completion_ui._get_total_tasks()
        self.assertEqual(total_tasks, 1)  # Only one non-comment task

    def test_get_current_task_number(self):
        """Test getting current task number."""
        current_task_number = self.task_completion_ui._get_current_task_number()
        self.assertEqual(current_task_number, 1)

if __name__ == '__main__':
    unittest.main()
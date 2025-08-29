import pytest
from src.Vector_Issue.utils import set_test_mode
from src.Vector_Issue import utils
@pytest.fixture
def reset_test_mode():
    """Fixture to reset the global test mode after each test."""
    yield
    set_test_mode(0)  # reset test mode after the test

class TestSetTestMode:
    def test_set_test_mode_enables_logging(reset_test_mode, capsys):
        """Test that enabling test mode allows logs to be printed."""
        test_mode = 1
        set_test_mode(test_mode)
        utils.test_log("Test message")

        captured = capsys.readouterr()
        assert "Test message" in captured.out

    def test_set_test_mode_disables_logging(reset_test_mode, capsys):
        """Test that disabling test mode prevents logs from being printed."""
        test_mode = 0
        set_test_mode(test_mode)
        utils.test_log("Test message")

        captured = capsys.readouterr()
        assert "Test message" not in captured.out

from python_rpa_example_bot import __version__
from python_rpa_example_bot.main import list_bucket_names

def test_version():
    assert __version__ == '0.1.0'

def test_list_bucket_names():
    assert list_bucket_names() == ["robocorp-test", "training-input-files", "training-output-folder"]
import pytest
import os
import shutil
from mailconverter.mailconverter import MailConverter
from pathlib import Path



def test_start():
    test_dir = "C:\\Users\\wob-admin\\repos\\MailConverter\\data"

    converter = MailConverter(debug=True)
    converter.start()
    # Check if the mht directory is created
    mht_dir = test_dir + "\\mht"

    mht_dir = Path(mht_dir)
    assert mht_dir.exists()
    # Check if the .mht file is created
    mht_dir = str(mht_dir)
    for i in range(1, 2):
        mht_file = mht_dir + "\\test" + str(i) + ".mht"
        mht_file = Path(mht_file)
        assert mht_file.exists()
        mht_file = str(mht_file)

    # Clean up by deleting the mht directory
    shutil.rmtree(mht_dir)


def test_gui():
    '''
    test the gui for this tool
    '''

    converter = MailConverter(interface="gui")
    converter.start()

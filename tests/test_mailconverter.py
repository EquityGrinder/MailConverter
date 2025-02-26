import pytest
import os
import shutil
from mailconverter.mailconverter import MailConverter


def test_start():
    '''
    Tests if the MailConverter class can be instantiated and creates the necessary directories
    and files.
    '''    
    
    converter = MailConverter(debug=True)
    converter.start()
    print("Testing the start method \n")

    test_files = ['test' + str(i) for i in range(1, 2)]
    print("test_files: ", test_files)
    test_dir = 'data/'  
    print("test_dir: ", test_dir)
    mht_dir = test_dir + "mht/"


    
    assert mht_dir.exists()

    # Check if the .mht file is created    
    for file in test_files:
        mht_file = mht_dir + file + ".mht"
        assert mht_file.exists()

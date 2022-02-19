import os
import argparse
#===============================================================================
def is_filename(filename:str):
    try:
        f = open(filename)
    except FileNotFoundError:
        raise argparse.ArgumentTypeError(f"Cannot open file'{filename}'")
    finally:
        try:
            f.close()
        except Exception:
            pass
    return filename
#===============================================================================
def is_directory(dir:str):
    if not os.path.isdir(dir):
        raise argparse.ArgumentTypeError(f"'{dir}' is not a valid directory")
    return os.path.abspath(dir)
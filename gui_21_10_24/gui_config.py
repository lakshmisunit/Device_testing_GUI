import os, sys
import argparse
HOME_PATH = os.path.expanduser("~")
Topic_Names = ['LED_GLOW', 'LED_GLOW1', 'LED_GLOW2']
Download_path = os.path.join(HOME_PATH, "Downloads")
mac_address_in_page = 70

def parse_arguments():
    parser = argparse.ArgumentParser(description='Device Tester')
    parser.add_argument('-n', '--Mac_in_page', type=str, default=mac_address_in_page, help='No of Mac addresses in a Page')
    return parser.parse_args()

args = parse_arguments()
rows_in_page = args.Mac_in_page

print(Download_path)

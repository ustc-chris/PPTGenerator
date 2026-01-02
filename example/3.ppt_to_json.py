import os, sys
sys.path.append("..")
from share.PPTGenerator import PPTReader
import json


output_json = "./output/convert_out.json"
input_ppt = "./output/test.pptx"

def main():
    ppt = PPTReader(input_ppt)
    data = ppt.convert_back()
    with open(output_json, 'w') as file:
        json.dump(data, file, ensure_ascii=False, indent=4)

if __name__ == "__main__":
    main()
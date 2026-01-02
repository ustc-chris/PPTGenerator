import os, sys
sys.path.append("..")
from share.PPTGenerator import PPT


input_json = "./output/All.json"
template_file = "../share/template.pptx"
output_ppt = "./output/test.pptx"

def main():
    ppt = PPT(template_file)
    ppt.read(json_file=input_json)
    ppt.write(output_ppt)

if __name__ == "__main__":
    main()
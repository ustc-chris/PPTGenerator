import os, sys
import json

thisfiledir = os.path.dirname(os.path.abspath(__file__))
txtfile = os.path.join(thisfiledir, 'All.txt')
txtoutfile = os.path.join(thisfiledir, 'output/All_process.txt')
jsonfile = os.path.join(thisfiledir, 'output/All.json')
pptfile = os.path.join(thisfiledir, 'output/SidePPT.pptx')
template_file = os.path.join(thisfiledir, '../share/template.pptx')


Titles = []


def main():
    if not os.path.exists(os.path.join(thisfiledir, 'output')):
        os.makedirs(os.path.join(thisfiledir, 'output'))
    Process_txt()
    Convert_json()
    print(Titles)
    Process_ppt()

def Process_txt():
    # 去掉所有连续>2行的空行
    cont_empty = 0
    with open(txtfile, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        with open(txtoutfile, 'w', encoding='utf-8') as fout:
            for line in lines:
                line = line.strip()
                if len(line) == 0:
                    cont_empty += 1
                else:
                    cont_empty = 0
                if cont_empty <= 2:
                    fout.write(line + "\n")
        
def Convert_json():
    PAGE = 1
    lyrics = {"slides":[]}
    lyrics["slides"].append({"template": "", "items": []})
    with open(txtoutfile, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        for line in lines:
            line = line.strip()
            # 标题行
            if line.startswith('#'):
                if lyrics["slides"][-1]["template"] == "":
                    lyrics["slides"][-1]["template"] = "Title_With_Author"
                else:
                    lyrics["slides"].append({
                        "template": "Title_With_Author",
                        "items": []
                    })
                line = line.strip("#")
                parts = line.split("#")
                if len(parts) == 2:
                    title = parts[0].strip()
                    author = parts[1].strip()
                else:
                    title = line.strip("#").strip()
                    author = ""
                lyrics["slides"][-1]["items"].append({"title": title, "author": author, "PAGE": PAGE})
                lyrics["slides"].append({"template": "", "items": []})
                PAGE += 1
                Titles.append([PAGE, title])
            # 空白页
            elif len(line) == 0:
                if lyrics["slides"][-1]["template"] == "Single" or lyrics["slides"][-1]["template"] == "Title_With_Author":
                    lyrics["slides"].append({"template": "", "items": []})
                elif lyrics["slides"][-1]["template"] == "":
                    lyrics["slides"][-1]["template"] = "Blank"
                    lyrics["slides"][-1]["items"].append({"PAGE": PAGE})
                    PAGE += 1
                    lyrics["slides"].append({"template": "", "items": []})
            else:
                if lyrics["slides"][-1]["template"] == "":
                    lyrics["slides"][-1]["template"] = "Single"
                    lyrics["slides"][-1]["items"].append({"text": line, "PAGE": PAGE})
                    PAGE += 1
                else:
                    lyrics["slides"][-1]["items"][-1]["text"] += "\n" + line
    with open(jsonfile, 'w', encoding='utf-8') as f:
        json.dump(lyrics, f, ensure_ascii=False, indent=4)


def Process_ppt():
    sys.path.append(os.path.join(thisfiledir, '..'))
    from share.PPTGenerator import PPT
    ppt = PPT(template_file)
    ppt.read(json_file=jsonfile)
    for title in Titles:
        ppt.add_section(title[0] - 1, title[1])
    ppt.write(pptfile)



if __name__ == '__main__':
    main()
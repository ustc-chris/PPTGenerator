
import json 
import os
input_txt = "./All.txt"
output_path = "./output"
output_json = "./output/All.json"

def main():
    # readlines
    input_lines = open(input_txt, "r").readlines()
    all_lines = ""
    for line in input_lines:
        all_lines += line
    all_lines = all_lines.split("\n\n")

    # process each line
    slides = {"slides": []}
    slides["slides"].append({"template": "EN_ZH", "items": []})
    for line in all_lines:
        line = line.strip()
        temp = line.split("\n")
        temp = [item.strip() for item in temp]
        if temp[0].startswith("=="):
            if not temp[-1].startswith("=="):
                print("ERROR!!!", temp)
            temp = temp[1:-1]
            length = len(temp)
            if length % 2 != 0:
                print("ERROR!!!", temp)
            length //= 2
            chinese = temp[0:length]
            chinese = "\n".join(chinese)
            english = temp[length:]
            english = "\n".join(english)
            # push previous
            slides["slides"].append({"template": "Poem_Translate", "items": []})
            # add new
            slides["slides"][-1]["items"].append({"text": chinese, "translate": english})
            # prepare next
            slides["slides"].append({"template": "EN_ZH", "items": []})
            continue
        if len(temp) == 3:
            if temp[0].startswith("ACT"):
                scene_num = temp[0].split(" ")[1].strip()
                scene_num_chinese = {"1": "一", "2": "二", "3": "三", "4": "四", "5": "五", "6": "六", "7": "七", "8": "八", "9": "九"}[scene_num]
                if len(slides["slides"]) > 1:
                    slides["slides"].append({"template": "Blank", "items": [{}]})
                    slides["slides"].append({"template": "Title_With_Author", "items": []})
                slides["slides"][-1]["template"] = "Title_With_Author"
                slides["slides"][-1]["items"].append(
                    {
                        "title": f"Scene {scene_num}\n{temp[2]}",
                        "author": f"第{scene_num_chinese}幕\n{temp[1]}"
                    }
                )
                slides["slides"].append({"template": "EN_ZH", "items": []})
            else:
                print("ERROR!!!", temp)
        elif len(temp) == 2:
            slides["slides"][-1]["items"].append({"en": temp[1], "zh": temp[0]})
        elif len(temp) == 4:
            slides["slides"][-1]["items"].append({"en": temp[3], "zh": temp[1]})
        else:
            print("ERROR!!!", temp)

    # json output
    if not os.path.exists(output_path):
        os.makedirs(output_path)
    with open(output_json, "w") as f:
        json.dump(slides, f, ensure_ascii=False, indent=4)


if __name__ == "__main__":
    main()
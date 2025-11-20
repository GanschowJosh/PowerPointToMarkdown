"""
takes every .pptx file in a given directory and outputs converted to md in a specified directory
"""
import os
from markitdown import MarkItDown
from pptx2md import convert, ConversionConfig
from pathlib import Path

def convertPPTX2MD(input_folder, output_folder):
    output_folder.mkdir(parents=True, exist_ok=True)
    
    md = MarkItDown()
    converted = False

    for path in input_folder.iterdir():
        if not path.is_file(): continue

        try:
            suffix = path.suffix.lower()
            output_path = output_folder / path.with_suffix(".md").name
            if suffix == ".pdf":
                result = md.convert(path)
                output_path.write_text(formatting(result.text_content), encoding='utf-8')
            elif suffix == '.pptx':
                image_dir = output_folder/'img'
                image_dir.mkdir(parents=True, exist_ok=True)
                convert(ConversionConfig(
                    pptx_path=path,
                    output_path=output_path,
                    image_dir=image_dir,
                ))
            converted = True
        except Exception as e:
            print(f"An error ({e}) occurred converting {path.name}")
            continue
    if converted: mergeMdFiles(output_folder)

def formatting(string):
    return string.replace(" __", "__").replace("\\-", "-").replace("\\.", ".").replace("\\,", ",").replace("\\(", "(").replace("\\)", ")").replace("\\#", "#").replace("\\+", "+").replace("\\!", "!").replace("\\[", "[").replace("\\]", "]").replace("\\_", "_").replace('\x0c', '')

def mergeMdFiles(directory):
    merged_content = ""
    
    for path in directory.glob("*.md"):
        if path.name == "Merged.md": continue

        text = path.read_text(encoding='utf-8')
        merged_content += formatting(text) + "\n\n"
    
    merged_path = directory / "Merged.md"
    merged_path.write_text(merged_content, encoding='utf-8')

if __name__ == "__main__":
    inputFolder = Path(input("Input folder:\t").strip())
    outputFolder = Path(input("Output folder:\t").strip())

    convertPPTX2MD(inputFolder, outputFolder)

"""
takes every .pptx file in a given directory and outputs converted to md in a specified directory
"""
import os
from pptx2md import convert, ConversionConfig

def convertPPTX2MD(inputFolder, outputFolder):
    if not os.path.exists(outputFolder):
        os.makedirs(outputFolder)

    for filename in os.listdir(inputFolder):
        try:
            if filename.endswith(".pptx"):
                inputPath = os.path.join(inputFolder, filename)
                outputPath = os.path.join(outputFolder, os.path.splitext(filename)[0] + ".md")
                curr_conversion_config = ConversionConfig(pptx_path=inputPath, output_path=outputPath, image_dir=f"{outputFolder}/img", disable_notes=True)
                convert(curr_conversion_config)
        except Exception as e:
            print(f"An error ({e}) occurred converting {filename}")
            continue
    mergeMdFiles(outputFolder)

def formatting(string):
    return string.replace(" __", "__").replace("\\-", "-").replace("\\.", ".").replace("\\,", ",").replace("\\(", "(").replace("\\)", ")").replace("\\#", "#").replace("\\+", "+").replace("\\!", "!").replace("\\[", "[").replace("\\]", "]").replace("\\_", "_")

def mergeMdFiles(directory):
    mergedContent = ""
    
    # Iterate over all files in the directory
    for filename in os.listdir(directory):
        # Check if the file is a .md file
        if filename.endswith(".md"):
            # Open the .md file
            with open(os.path.join(directory, filename), 'r', encoding='utf-8') as file:
                # Read the content and add it to the merged content
                mergedContent += formatting(file.read()) + "\n\n"
                
    # Write the merged content to a new .md file
    with open(os.path.join(directory, "Merged.md"), 'w', encoding='utf-8') as file:
        file.write(mergedContent)

if __name__ == "__main__":
    inputFolder = input("Input folder:\t").replace("\\", "/") 
    outputFolder = input("Output folder:\t").replace("\\", "/")

    convertPPTX2MD(inputFolder, outputFolder)

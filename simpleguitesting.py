import os
import PySimpleGUI as sg
from pptx2md import convert

def convertPPTX2MD(inputFolder, outputFolder):
    if not os.path.exists(outputFolder):
        os.makedirs(outputFolder)

    for filename in os.listdir(inputFolder):
        try:
            if filename.endswith(".pptx"):
                inputPath = os.path.join(inputFolder, filename)
                outputPath = os.path.join(outputFolder, os.path.splitext(filename)[0] + ".md")

                convert(pptx_path=inputPath, output=outputPath)
        except Exception as e:
            print(f"An error ({e}) occurred converting {filename}")
            continue
    mergeMdFiles(outputFolder)

def formatting(string):
    return string.replace(" __", "__").replace("\\-", "-").replace("\\.", ".").replace("\\,", ",").replace("\\(", "(").replace("\\)", ")").replace("\\#", "#").replace("\\+", "+").replace("\\!", "!").replace("\\[", "[").replace("\\]", "]").replace("\\_", "_")

def mergeMdFiles(outputFolder):
    mergedContent = ""

    for filename in os.listdir(outputFolder):
        if filename.endswith(".md"):
            with open(os.path.join(outputFolder, filename), 'r', encoding='utf-8') as file:
                mergedContent += formatting(file.read()) + "\n\n"

    with open(os.path.join(outputFolder, "Merged.md"), 'w', encoding='utf-8') as file:
        file.write(mergedContent)

def main():
    layout = [
        [sg.Text('Input Folder:'), sg.InputText(key='input_text', size=(30, 1), enable_events=True), sg.FolderBrowse()],
        [sg.Text('Output Folder:'), sg.InputText(key='output_text', size=(30, 1), enable_events=True), sg.FolderBrowse()],
        [sg.Button('Submit')],
        [sg.Multiline('', size=(50, 10), key='log_pane')],
    ]

    window = sg.Window('PowerPoint to Markdown Converter', layout, finalize=True)

    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED:
            break
        elif event == 'Submit':
            input_folder = values['input_text']
            output_folder = values['output_text']
            
            if input_folder and output_folder:
                convertPPTX2MD(input_folder, output_folder)
                print('Conversion completed successfully!')
                mergeMdFiles(output_folder)
                print('Merge completed successfully!')
                window['input_text'].update('')  # Clear the input text box
                window['output_text'].update('')  # Clear the output text box
                window['log_pane'].print(input_folder)
                
    window.close()

if __name__ == '__main__':
    main()

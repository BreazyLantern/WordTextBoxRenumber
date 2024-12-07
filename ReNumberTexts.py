from pathlib import Path
import os
from datetime import datetime
import time
import win32com.client


current_dir = Path(__file__).parent.parent if "__file__" in locals() else Path.cwd()

today_date = datetime.today().strftime("%d %b, %Y")

def get_doc(file, date):
    #open up a invisible word doc
    word_app = win32com.client.DispatchEx("Word.Application")
    word_app.Visible = False
    word_app.DisplayAlerts = False

    #temp file name
    result = file

    word_app.Documents.Open(str(file))

    # Loop through all the shapes
    for i in range(word_app.ActiveDocument.Shapes.Count):
        if word_app.ActiveDocument.Shapes(i + 1).TextFrame.HasText:
            words = word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words
            # add 6 to the texts
            addsix = int(words.Item(1).Text) + 6
            words.Item(1).Text = addsix

    #    Save the new file
    curr_time = time.strftime("%H_%M_%S", time.localtime())
    formatedDateTime = f"_{date}_time_{curr_time}"
    if file.stem.find("_") != -1:
        index = file.stem.find("_")
        #print(index)
        replace = file.stem[index:]
        #print(replace)

        result = file.stem.replace(replace, formatedDateTime) + file.suffix
    else:
        result = file.stem + formatedDateTime + file.suffix
    #print(result)
    output_path = current_dir / f"{result}"
    word_app.ActiveDocument.SaveAs(str(output_path))
    word_app.ActiveDocument.Close(SaveChanges=False)
    word_app.Application.Quit()
    print("Execution Completed")


def exec():
    doc = str(input("Please type in the name of the document: "))

    file_path = os.path.abspath(current_dir / doc)
    if os.path.exists(file_path):
        path = current_dir / doc
        print(str(path))
        #for doc_file in Path(current_dir).rglob(doc):
            # Open each document and replace string
        get_doc(path, today_date)
    else:
        print(f"{current_dir / doc} did not exist")

done = True
while(done):
    exec()
    if input("type 'no' to stop program").lower() == "no":
        done = False
import json
import pandas
import pg8000
import openpyxl
import sys
from datetime import datetime
from os import path, listdir, sep, makedirs

def getValidFilename(name):
    return "".join([x.lower() if x.isalnum() else "_" for x in name])

lenArgv = len(sys.argv)
makedirs("query", exist_ok=True)
makedirs("output", exist_ok=True)

if (lenArgv < 1):
    sys.exit()
elif (lenArgv > 1):
    sys.argv[1] = sys.argv[1].lower()

if (lenArgv == 1 or sys.argv[1] == "help"):
    print("""
    python run.py query sql_file_path|sql_dir_path [excel_file_name]
    python run.py help
    """)
elif (sys.argv[1] != "query"):
    print("command not found")
elif (sys.argv[1] == "query" and lenArgv == 2):
    print("file|dir not found")
else:
    settingJson = open("setting.json", "r").read()
    setting = json.loads(settingJson)
    connection = pg8000.connect(**setting)

    sqlFilenames = []
    sheetNames = []
    fileName = "output" + sep + datetime.now().strftime("%Y%m%d_%H%M") + "_"
    if (path.isdir(sys.argv[2])):
        fileName += getValidFilename(sys.argv[2]) + ".xlsx"
        for sqlFilename in sorted(listdir(sys.argv[2])):
            sqlFilepath = path.join(sys.argv[2], sqlFilename)
            if (path.isfile(sqlFilepath)):
                sqlFilenames.append(sqlFilepath)
                sheetNames.append(getValidFilename(path.splitext(path.basename(sqlFilename))[0]))
    else:
        fileName += getValidFilename(path.splitext(path.basename(sys.argv[2]))[0]) + ".xlsx"
        sqlFilenames.append(sys.argv[2])
        sheetNames.append("sheet")

    try:
        writer = pandas.ExcelWriter(fileName, engine='openpyxl')
        for sqlFilename, sheetName in zip(sqlFilenames, sheetNames):
            query = open(sqlFilename, "r").read()
            data = pandas.read_sql(query.replace("%", "%%"), connection)
            data.to_excel(writer, sheet_name=sheetName, index=False, encoding="utf-8")
        writer.save()
        writer.close()
    except Exception as ex:
        print(ex)
        try:
            writer.close()
        except:
            pass

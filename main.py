from app.definitions import *

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

if not os.path.exists("./main_res"):
    os.makedirs("./main_res")

while True:
    if not os.path.exists("./src"):
        os.makedirs("./src")

    files = []
    i = 0
    for excel in os.listdir("./src"):
        if excel.endswith(".xlsx"):
            files.append(excel)
            print(i, "-", excel)
            i += 1

    if len(files) == 0:
        print("No such data found in src directory.")
        exit()

    # Get the location of the data
    try:
        file_key = input("\nInput file number: ")
        file_name = files[int(file_key)]
        path = "src/" + file_name
        print("Excel File: " + file_name)
    except Exception as e:
        print("Error: ", str(e))

    # Read the data
    try:
        data = pd.read_excel(path, sheet_name=None, index_col=None, header=None)

        # Get the keys
        xl = pd.ExcelFile(path)
        keys = xl.sheet_names

        # Iterate through the sheets
        for key in keys:
            if key not in notIncluded:
                print(key)

                # Vessel Name
                vessel = data[key].iloc[0, 2]

                # Machinery Name
                machinery = data[key].iloc[2, 2]

                # Start traversing the data on row 7
                row = 7
                isValid = True

                # Prepare the sheets
                book = Workbook()
                sheet = book.active

                sheet.append(header)

                while isValid:

                    rowData = (
                        vessel,
                        machinery,
                    )

                    for col in range(7):
                        d = data[key].iloc[row, col]

                        if (pd.isna(d)) and (col == 0):
                            isValid = False
                            break

                        if pd.isna(d):
                            d = " "

                        if ((col == 4) or (col == 5)) and isinstance(d, datetime):
                            d = d.strftime("%d-%b-%y")
                        else:
                            d = re.sub("\\s+", " ", str(d))

                        tempTuple = (d,)
                        rowData += tempTuple

                    if isValid:
                        sheet.append(rowData)
                        row += 1

                create_name = file_name[: len(file_name) - 4]
                creation_folder = "./main_res/" + create_name
                if not os.path.exists(creation_folder):
                    os.makedirs(creation_folder)
                book.save(creation_folder + "/" + key + ".xlsx")

        print("Done...")
    except Exception as e:
        print("Error: " + str(e))

    isContinue = input("Input 1 to continue: ")
    if isContinue != "1":
        break

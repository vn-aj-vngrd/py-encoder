from app.definitions import *

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

if not os.path.exists("./main_res"):
    os.makedirs("./main_res")

main_function()
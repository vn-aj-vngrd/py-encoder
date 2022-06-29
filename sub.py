from app.definitions import *

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

if not os.path.exists("./sub_res"):
    os.makedirs("./sub_res")

sub_function()
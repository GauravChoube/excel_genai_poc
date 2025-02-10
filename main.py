from excelParsing import ExcelVBAProcessor
from openaiClient import OPENAI_CLIENT
import json

CONFIG_PATH = "config.json"
EXCEL_PATH = r"C:\Users\gaurav.j.choubey\Desktop\project\gtic-7-2025\excelsheet_poc\excel_genai_poc\examplemacro1.xlsm"

def readConfig():
    with open(CONFIG_PATH, 'r') as file:
        config = json.load(file)
    print(config)

    if len(config)==0:
        raise(f"ERROR:Invalid config file.Please provide openAI creadential")

    elif 'openaiKey' not in config or len(config['openaiKey'])==0:
        raise(f"ERROR:Invalid or missing openaiKey .Please provide 'openaiKey'")
    
    elif 'openaiEndPoint' not in config or len(config['openaiEndPoint'])==0:
        raise(f"ERROR:Invalid or missing openaiKey .Please provide 'openaiEndPoint'")

    elif 'openaiVersion' not in config or len(config['openaiVersion'])==0:
        raise(f"ERROR:Invalid or missing openaiKey .Please provide 'openaiVersion'")

    return config

def main():
    try:
        pass
        # read the config file
        config = readConfig()

        # create openai client
        openaiClient = OPENAI_CLIENT(azure_endpoint=config['openaiEndPoint'],api_key=config['openaiKey'],api_version=config['openaiVersion'],model=config['model'])
        # res = openaiClient.promptCall(prompt="Hi, Are you able to convert codes")
        # print(res)

        # create POC object
        processor = ExcelVBAProcessor(file_path=EXCEL_PATH,openAIClient=openaiClient)
        processor.process_excel_file()

        # print("Running converted Python class...")

    except Exception as msg:
        print(f"ERROR: Exception with message : '{msg}'")

if __name__ == '__main__':
    main()
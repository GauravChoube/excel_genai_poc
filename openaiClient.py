class OPENAI_CLIENT:
    def __init__(self,apiKey,endPoint):
        pass

    def setup():
        pass
        
    def call(promt,code):
        pass
        

        
    def test():
        pass

if __name__ == '__main__':
    obj = OPENAI_CLIENT()
    obj.call()


import os
import win32com.client
from openpyxl import load_workbook
from openai import AzureOpenAI

class OPENAI_CLIENT:
    def __init__(self, azure_endpoint, api_key, api_version,model):
        if not azure_endpoint or not api_key:
            raise ValueError("Azure OpenAI credentials are missing!")
        self.azure_endpoint = azure_endpoint
        self.api_key        = api_key
        self.api_version    = api_version
        self.model          = model

        self.client = AzureOpenAI(
            azure_endpoint=azure_endpoint, 
            api_key=api_key,               
            api_version=api_version
        )
        

        print(f"connected to openai for model '{model}':{self.client.is_closed()}")
        

    def promptCall(self, prompt):
        messages = [
            {"role": "system", "content": "You are an program converter that converts VBA script to Python script."},
            {"role": "user", "content": prompt}
        ]
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=messages
        )
        return response.choices[0].message.content

    @classmethod
    def connect(cls):
        pass

    def convert_vba_to_python(self, vba_code):
        prompt = f"Convert the following VBA script enclosed in single inverted comas to Python script and also take care of dependencies in VBA script and Dont add any commets and description about logic and example. Just provide clean python function:\n\n'{vba_code}'"
        return self.promptCall(prompt)


def extract_vba_from_excel(file_path):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  
    
    workbook = excel.Workbooks.Open(file_path)
    vba_code_modules = {}

    try:
        for module in workbook.VBProject.VBComponents:
            if module.Type == 1:  
                module_name = module.Name
                vba_code = module.CodeModule.Lines(1, module.CodeModule.CountOfLines)
                vba_code_modules[module_name] = vba_code
    except Exception as e:
        print("Error extracting VBA:", e)
    finally:
        workbook.Close(SaveChanges=False)
        excel.Quit()

    return vba_code_modules
 

# def main():
#     file_path = "C:\\Users\\b.sowkoor.manju.raju\\Downloads\\examplemacro1.xlsm"  
#     azure_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT") 
#     api_key = os.getenv("AZURE_OPENAI_API_KEY")         

#     if not azure_endpoint or not api_key:
#         raise ValueError("Azure OpenAI credentials are missing!")

#     obj = OA_CLIENT(azure_endpoint, api_key)
    
#     vba_modules = extract_vba_from_excel(file_path)

#     converted_python_code = []
    
#     for module_name, vba_code in vba_modules.items():
#         print(f"Converting module: {module_name}...")
#         python_code = obj.convert_vba_to_python(vba_code)
#         converted_python_code.append(f"# Converted from {module_name}\n{python_code}\n")

#     # Save to a Python file
#     output_file = "converted_script.py"
#     with open(output_file, "w", encoding="utf-8") as f:
#         f.write("\n".join(converted_python_code))

#     print(f"Conversion complete. Saved to {output_file}")


# if __name__ == "__main__":
#     main()

def main():
    file_path = "C:\\Users\\b.sowkoor.manju.raju\\Downloads\\examplemacro1.xlsm"  
    azure_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT") 
    api_key = os.getenv("AZURE_OPENAI_API_KEY")         

    if not azure_endpoint or not api_key:
        raise ValueError("Azure OpenAI credentials are missing!")

    obj = OPENAI_CLIENT(azure_endpoint, api_key)
    
    vba_modules = extract_vba_from_excel(file_path)

    for module_name, vba_code in vba_modules.items():
        print(f"\n=== Converted Python Code from {module_name} ===\n")
        python_code = obj.convert_vba_to_python(vba_code)
        print(python_code)  
        print("\n" + "="*50 + "\n") 

if __name__ == "__main__":
    main()
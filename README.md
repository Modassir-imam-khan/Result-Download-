# jee_results_collection

## Steps to Run:
- Download and extract the code
- Navigate to the code folder and run:  `pip install -r requirements.txt`
- Update the the input file
- For the first time, run:  
  - Mains: `python main.py --input="inputs.xlsx" --output="results.xlsx" --mode="start over" --img_path="images" --exam="Mains"` 
  - Advanced: `python main.py --input="inputs-advanced.xlsx" --output="results-advanced.xlsx" --mode="start over" --img_path="images-advanced" --exam="Advanced"`
  - CBSE Result: `python main.py --input="input-cbse-2020.xlsx" --output="result-cbse-2020.csv" --mode="start over" --img_path="images-cbse" --exam="CBSE"`
  - Maharashtra result: `python main.py --input="input-maharashtra-2020.xlsx" --output="result-maharashtra-2020.csv" --mode="start over" --img_path="images-maharashtra" --exam="MAHARASHTRA"`
  - KVPY Result: `python main.py --input="input-kvpy.xlsx" --output="result-kvpy.csv" --mode="start over" --img_path="images-KVPY" --exam="KVPY"`
  - NEET - Confimation page data: `python main.py --input="inputs-neet-conf.xlsx" --output="results-neet-conf.xlsx" --mode="start over" --img_path="images-neet-conf" --exam="NEET-CONF" --full_screenshot="True"`
  - NEET - Admit card data: `python main.py --input="inputs-neet-admit.xlsx" --output="inputs-neet.xlsx" --mode="start over" --img_path="images-neet-admit" --exam="NEET-ADMIT" --full_screenshot="True"` **Note: Refresh the captcha before entering, and full-screenshot = "True" is to take the screenshot of the entire page. Otherwise it takes the screenshot of whatever is visible on without scrolling, the script runs without providing this as by default it is assumed to be "False"**
  - NEET: `python main.py --input="inputs-neet.xlsx" --output="results-neet.csv" --mode="start over" --img_path="images-neet" --exam="NEET"`
  - NEET - OMR data: `python main.py --input="inputs-neet-conf.xlsx" --output="results-neet-omr.xlsx" --mode="start over" --img_path="images-neet-omr" --exam="NEET-OMR"` **Note: the file 'results-neet-omr.xlsx' is a dummy file and will contain no useful information**
  - WBJEE - Both conf page and rank card: `python main.py --input="inputs-wbjee.xlsx" --output="results-wbjee.csv" --mode="start over" --img_path="images-wbjee" --exam="WBJEE"` **Note: the confimation pdf will be inside the img_path only, along with the screenshots**
  - COMEDK - Both app pdf and rank card pdf: `python main.py --input="inputs-comedk.csv" --output="results-comedk.csv" --mode="start over" --img_path="images-comedk" --exam="COMEDK"` **Note: both pdfs will be inside the img_path, along with dummy screenshots (ignore these, they are for bookkeeping purposes). The results are not perfect due to inconsistent pdf formatting, so some manual post-processing may be required.**

- To continue the collection from where it was stopped/crashed in the last run, change the --mode to 'continue' from 'start over'
- Filenames ending with ".xlsx" will be saved as excel file, while filenames ending with ".csv" will be saved as csv files. 

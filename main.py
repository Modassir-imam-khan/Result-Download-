from fastprogress import progress_bar
import pandas as pd
import numpy as np
import sys, os
import shutil
from helium import *
import argparse
import time
from tqdm import tqdm
from result_utils import *

def get_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--input', help='input file ', default='inputs.xlsx')
    parser.add_argument('-o', '--output', help='output file', default='results.xlsx')
    parser.add_argument('-m', '--mode', help='mode', default='continue')
    parser.add_argument('-img', '--img_path', help='image directory', default='images')
    parser.add_argument('-e', '--exam', help='which exam Mains/Advanced', default='Mains')
    parser.add_argument('-fs', '--full_screenshot', default="False")
    args = parser.parse_args()
    return args 


if __name__ == "__main__":
    args = get_args()
    input_file = args.input
    results_file = args.output
    mode = args.mode
    img_path = Path(args.img_path)
    exam = args.exam


    df = read_df(input_file)
    if "Date of Birth" in df.columns:
        df['Date of Birth'] = pd.to_datetime(df['Date of Birth'], dayfirst=True).dt.strftime('%d-%m-%Y')  

    if not Path(results_file).exists() or mode=='start over': 
        save_df(results_file, pd.DataFrame())
        if img_path.exists(): 
            shutil.rmtree(str(img_path))
            
    rdf = read_df(results_file) 
    img_path.mkdir(exist_ok=True)

    secs = np.arange(0.1, 5, 0.01)
    get_data_funcs = {'Mains': get_data_mains, 'Advanced': get_data_advanced, 
                      'NEET': get_data_neet, 'NEET-ADMIT': get_neet_admit_data,
                      'NEET-CONF': get_neet_conf_data, 'NEET-OMR': get_neet_omr_data,
                      'WBJEE': get_wbjee_data, 'COMEDK': get_comedk_data,
                      'CBSE': get_cbse_data, 'MAINS-2021': fill_mains_details,
                      'MAHARASHTRA': get_maharashtra_data,'KVPY':get_kvpy_result } #edited
    get_data = get_data_funcs[exam]

    for i in progress_bar(range(len(df))):
        name = get_prop(df, i, "Name")
        dob = get_prop(df, i, 'Date of Birth')
        app_no = get_prop(df, i, 'Application No.')
        roll_no = get_prop(df, i, 'Roll Number')
        batch = get_prop(df, i, 'Batch')
        phone = get_prop(df, i, 'Phone No.')
        
        
        img_file = img_path/f"{app_no}_{name}.png"

        if mode=='start over' or img_file.exists()==False:
            data, browser = get_data(app_no=app_no, dob=dob, phone=phone, roll_no=roll_no, df=df, i=i, 
                                     img_path=img_path, name=name)    
            for c, v in [("Candidate's Name", name), ('Date of Birth', dob), 
                         ('Application No.', app_no), ('Batch', batch)]:
                if c not in list(data.keys()) and str(v) != "nan":
                       data[c] = v
            rdf = pd.concat([rdf, pd.DataFrame(data, index=[0])])
            save_df(results_file, rdf)
            
            save_screenshot(browser, img_file, fullpage=eval(args.full_screenshot))
            kill_browser()

            #Random wait
            time.sleep(np.random.choice(secs, 1)[0])
# pip install python-docx
# pip install openpyxl
# pip install pandas
# pip install regex


# pip install pyinstaller
# pyinstaller --onefile --add-data "input;input" --add-data "output;output" main.py 

import pandas as pd
import os
import re
from docx import Document

def main(input_folder,output_folder):


    
    # Initialize lists to store individual pieces of information
    names = []
    address_list = []
    all_countries = []
    EMAIL = []
    PHONE = []
    timeList = []
    resortList = [] 
    origin_site = []
    other = []

    # List the contents of the directory
    file_list = os.listdir(input_folder)

    # Print the list of files
    print("Number of files:", len(file_list))

    for file_name in file_list:
        file_path = os.path.join(input_folder, file_name)
        if os.path.isfile(file_path) and file_name.endswith('.docx'):
            # Create a new Document object
            doc = Document(file_path)

            # Initialize an empty string to store the text
            text = ""

            # Iterate through paragraphs and append text to the string
            for paragraph in doc.paragraphs:
                text += paragraph.text + "\n"

            # Split the data into individual records based on blank lines
            records = text.split('\n\n')

            # Process each record and extract information
            for record in records:
                
                # Split each record into lines
                lines = record.split('\n')
                print(lines)

                # name
                names.append(lines[0])


                #address

                address_pattern = re.compile(r'(\d+\s+[^\n]+)\n([^\n]+,\s*[A-Z]{2}\s+\d{5}|NO ADDRESS ON FILE)')

                matches = re.search(address_pattern, record)
                if matches:
                    street, city_state_zip = matches.groups()
                    address = f"{street}, {city_state_zip}"
                    address_list.append(address)
                else:
                    address_list.append("NA")


                # country
                country_pattern = re.compile(r'UNITED STATES')
                country_match = country_pattern.search(record)
                # print("all_countries",len(match))
                if country_match:
                    country = country_match.group()
                    all_countries.append(country)

                else:
                    all_countries.append("NA")
        
                #email

                email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
                email_matches = re.findall(email_pattern, record)

                if email_matches:
                    email = email_matches
                    str_email = "".join(email).replace('[','').replace(']','')
                    
                    EMAIL.append(str_email)
                else:
                    email = "NA"
                    EMAIL.append(str_email)

                #phone number
                phone_pattern = re.compile(r'\b(\d{3}-\d{3}-\d{4})\b')

                # Process each record and extract phone numbers
                match = phone_pattern.search(record)
                if match:
                    phone = match.group(1)
                    PHONE.append(phone)
                else:
                    PHONE.append("NA")

                
                #time
                time_pattern = re.compile(r'Best Time To Call: (\d{1,2}:\d{2}[APMapm]{2})')

                matches = re.findall(time_pattern, record)
                if matches:
                    time = matches[0]
                    timeList.append(time)
                else:
                    time = "NA"
                    timeList.append(time)


                #resort
                record_str = ' '.join(lines)

                # Define the regex pattern to find the "Best Time To Call" statement and the data next to it
                best_time_to_call_pattern = re.compile(r'Best Time To Call: \d{1,2}:\d{2}[APMapm]{2}\s*([A-Z\s]+)')
                phone_number_pattern = re.compile(r'\d{3}-\d{3}-\d{4} \(Primary Phone #, [A-Z]+\)\s*([A-Z\s]+)')

                # Find the match in the string
                match1 = best_time_to_call_pattern.search(record_str)
                match2 = phone_number_pattern.search(record_str)

                if match1 :
                    resort = match1.group(1).strip()
                    resortList.append(resort)

                elif match2:
                    resort = match2.group(1).strip()
                    resortList.append(resort)

                else:
                    resort = "NA"
                    resortList.append(resort)
                
                #origin site
                pattern = r'\(ORIG\. SITE: ([A-Za-z0-9.\-_\s]+)\)'

                matches = re.search(pattern, record)
                if matches:
                    orig_site = matches.group(1)
                    origin_site.append(orig_site)
                else:
                    orig_site = "NA"
                    origin_site.append(orig_site)


                #other
                record_str = ' '.join(lines)
                pattern = r'\(ORIG\. SITE: ([A-Za-z0-9.\-_\s]+)\)'

                matches = re.finditer(pattern, record_str)
                found_match = False

                for match in matches:
                    found_match = True
                    pattern_end = match.end()  # Find the end position of the matched pattern
                    next_element_match = re.search(r'(.*?)(?=\(ORIG\. SITE:|$)', record_str[pattern_end:], re.DOTALL)  # Find the next element with spaces
                    if next_element_match:
                        next_element = next_element_match.group(1).strip()
                        other.append(next_element)
                    else:
                        next_element = "NA"
                        other.append(next_element)

                if not found_match:
                    other.append("NA")





                
        
    print("names:", len(names))
    print("address :",len(address_list))
    print("country :",len(all_countries))
    print("email",len(EMAIL))
    print("phone",len(PHONE))
    print("time :",len(timeList))
    print("resorts:", len(resortList))
    print("originSite:",len(origin_site))
    print("other:",len(other))
    dataList = []

    for name_ ,address_,country_ ,email_ ,phone_,best_time_to_call_,resorts_,origin_site_,other_ in zip(names,address_list,all_countries,EMAIL,PHONE,timeList,resortList,origin_site,other):
        dataList.append([name_ ,address_, country_ ,email_,phone_,best_time_to_call_,resorts_,origin_site_,other_])
    df = pd.DataFrame(dataList,columns=["name_" ,"address_","country_","email_","phone_","best_time_to_call_","resorts_","origin_site_","other_"])
    # output_folder = os.path.join("output", "LEADS.xlsx")
    
    print(df)
    counter = 1

    # Generate the base file name
    base_filename = "LEADS"

    # Generate the output file name with the counter
    output_file = os.path.join(output_folder, f"{base_filename}{counter}.xlsx")

    # Check if the file already exists, and increment the counter if needed
    while os.path.exists(output_file):
        counter += 1
        output_file = os.path.join(output_folder, f"{base_filename}{counter}.xlsx")

    # Save the DataFrame to an Excel file with the generated name using 'openpyxl' engine
    df.to_excel(output_file, index=False, engine='openpyxl')

    print(f"DataFrame saved to {output_file}")


main(r'input', r'output')
               
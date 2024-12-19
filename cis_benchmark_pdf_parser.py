import os
import re
import json
import pandas as pd
import fitz  # PyMuPDF

# Directory containing the PDF files
directory = "."  # current directory

# Loop through each PDF file in the current directory
for filename in os.listdir(directory):
    if filename.endswith(".pdf"):
        pdf_path = os.path.join(directory, filename)
        base_filename = os.path.splitext(filename)[0]  # Get the filename without extension
        
        # Define output file paths based on the current PDF file name
        cistext = f"{base_filename}.txt"
        cisjson = f"{base_filename}.json"
        cisexcel = f"{base_filename}.xlsx"

        # Extract text from the PDF using PyMuPDF
        print(f"[+] Converting '{filename}' to text...")
        pdf_text = ""
        
        with fitz.open(pdf_path) as pdf_file:
            for page_num in range(pdf_file.page_count):
                page = pdf_file[page_num]
                pdf_text += page.get_text()
        
        # Write the extracted text to a text file
        print(f"[+] Creating text file '{cistext}'...")
        with open(cistext, 'w', encoding='utf-8') as f:
            f.write(pdf_text)

        # Remove blank lines
        temp_text_file = f"{base_filename}_temp.txt"
        with open(cistext, 'r', encoding='utf-8') as filer:
            with open(temp_text_file, 'w', encoding='utf-8') as filew:
                for line in filer:
                    if line.strip():  # Only write non-blank lines
                        filew.write(line)

        # JSON Parsing Logic (Replace with your actual JSON parsing logic as needed)
        print(f"[+] Converting '{filename}' to JSON...")
        flagStart, flagComplete, flagSkipToPageBreak = False, False, False
        flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs, flagControl, flagMetre = False, False, False, False, False, False, False, False, False, False, False
        cis_name, cis_level, cis_desc, cis_ration, cis_impact, cis_audit, cis_remed, cis_defval, cis_refs, cis_control, cis_metre = "","","","","","","","","", "", ""
        listObj = []
        
        with open(cistext, 'r', encoding='utf-8') as filer:
            for line in filer:
                
                x = {} #json object
                
				
                if re.match(r"^(\d{1,2}(\.\d{1,2}){2,4})", line):    # identified CIS item name	
                    #print(line)
                    if cis_name != "" and cis_desc:
                        cis_name = cis_name.replace(' \n\n','').replace('\n\n','').rstrip()
                        cis_desc = cis_desc.replace(' \n\n','XMXM').replace('\n','').replace('XMXM','\n\n').rstrip()
                        cis_ration = cis_ration.replace(' \n\n','XMXM').replace('\n','').replace('XMXM','\n\n').rstrip()
                        cis_impact = cis_impact.replace(' \n\n','XMXM').replace('\n','').replace('XMXM','\n\n').rstrip()
                        cis_audit = cis_audit.replace(' \n\n','XMXM').replace('\n','').replace('XMXM','\n\n').rstrip()
                        cis_defval = cis_defval.replace(' \n\n','XMXM').replace('\n','').replace('XMXM','\n\n').rstrip()
                        cis_refs = cis_refs.rstrip()
                        cis_refs = cis_refs.strip().split('\n')
                        cis_refs = [item.strip() for item in cis_refs if item]
                        cis_level = cis_level.split('\n')
                        cis_level = [item.strip() for item in cis_level if item]
                        cis_control = cis_control.replace(' \n\n','XMXM').replace('\n','').replace('XMXM','\n\n').rstrip()
                        cis_metre = cis_metre.replace(' \n\n','XMXM').replace('\n','').replace('XMXM','\n\n').rstrip()
                        pattern = r"(#\s*.*|[A-Za-z]+\\[^\n]*)(?=\n|$)"

                        if "Appendix" in cis_control:
                            print(cis_control)
                            cis_control = cis_control[:cis_control.index("Appendix")]

                        if "Appendix" in cis_metre:
                            print(cis_metre)
                            cis_metre = cis_metre[:cis_metre.index("Appendix")]



                        # Search for the code block or configuration path in the text
                        match = re.search(pattern, cis_audit, re.DOTALL)
                        if match:
                            main_audit_text = cis_audit[:match.start()].strip()
                            code_block_text = cis_audit[match.start():].strip()
                        else:
                            main_audit_text = cis_audit
                            code_block_text = ""

                        match = re.search(pattern, cis_remed, re.DOTALL)
                        if match:
                            main_remed_text = cis_remed[:match.start()].strip()
                            remed_code_block_text = cis_remed[match.start():].strip()
                        else:
                            main_remed_text = cis_remed
                            remed_code_block_text = ""

                        match = re.match(r"^((\d+\.)+\d+)\s+(.*)", cis_name, re.DOTALL)
                        if match:
                            numeric_part = match.group(1)  # Extracts "3.3.7"
                            cleaned_title = match.group(3).strip()  # Extracts "Ensure Reverse Path Filtering is enabled"
                        else:
                            numeric_part = ""
                            cleaned_title = cis_name

                        pattern = r"\((.*?)\)"
                        match = re.search(pattern, cis_name)
                        if match:
                            # Extract the string inside the parentheses
                            method = match.group(1)
                            
                            # Remove the matched part from the original string (including the parentheses)
                            title = re.sub(pattern, '', cleaned_title).strip()

                        if not cis_control.strip():
                            result = []
                        else:
                            # Remove unwanted tokens (IG 1, IG 2, IG 3, and \u25cf)
                            cleaned_text = re.sub(r"(IG \d|\u25cf)", "", cis_control).strip()

                            # Regular expression to match Controls Version and Control text
                            pattern = r"(v\d+)\s+([\d.]+ [^\v]+?(?=(v\d+|$)))"  

                            # Find all matches
                            matches = re.findall(pattern, cleaned_text)

                            # Convert matches to the desired list of dictionaries
                            result = [{"Controls Version": match[0], "Control": match[1].strip()} for match in matches]

                        formatted_data = []
                        if cis_metre:
                            components = cis_metre.split()

                            # Initialize keys and data structure
                            keys = ["Techniques / Sub-techniques", "Tactics", "Mitigations"]
                            data = {key: [] for key in keys}

                            # Parse the components into respective categories
                            for component in components:
                                if component.startswith("T1"):
                                    data["Techniques / Sub-techniques"].append(component)
                                elif component.startswith("TA"):
                                    data["Tactics"].append(component)
                                elif component.startswith("M1") and not component.startswith("Techniques"):  # Exclude "Mitigations" label
                                    data["Mitigations"].append(component)

                            # Create the final formatted dictionary
                            formatted_data = [
                                {
                                    "Techniques / Sub-techniques": ", ".join(data["Techniques / Sub-techniques"]),
                                    "Tactics": ", ".join(data["Tactics"]),
                                    "Mitigations": ", ".join(data["Mitigations"]),
                                }
                            ]
                
                        


                        x['ID'] = numeric_part
                        x['Title'] = title
                        x['Method'] = method
                        x['Profile Applicability'] = cis_level
                        x['Description'] = cis_desc
                        x['Rationale'] = cis_ration
                        x['Impact'] = cis_impact
                        x['Audit'] = main_audit_text
                        x['Audit Commands'] = code_block_text
                        x['Remediations'] = main_remed_text
                        x['Remediation Commands'] = remed_code_block_text
                        x['Default value'] = cis_defval
                        x['References'] = cis_refs
                        x['CIS Controls'] = result
                        x['MITRE ATT&CK Mappings'] = formatted_data
                        # print(x)
                        cis_name, cis_level, cis_desc, cis_ration, cis_impact, cis_audit, cis_remed, cis_defval, cis_refs, cis_control, cis_metre = "","","","","","","","","","", ""
                        flagStart = False
                        # parsed = json.loads(x)
                        # print(json.dumps(x, indent=4))
                        listObj.append(x)

                    cis_name, cis_level, cis_desc, cis_ration, cis_impact, cis_audit, cis_remed, cis_defval, cis_refs, cis_control, cis_metre = "","","","","","","","","","", ""
                    flagStart = True
                    flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs, flagControl, flagMetre = True, False, False, False, False, False, False, False, False, False, False

                if flagStart:
                    if "Page" in line or "P a g e" in line:
                        continue
                    # from here we will be handling what section we're in and turning on and off the appropriate flags
                    if "Profile Applicability:" in line:	
                        flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs, flagControl, flagMetre = False, True, False, False, False, False, False, False, False, False, False
                    if "Description:" in line:	
                        flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs, flagControl, flagMetre = False, False, True, False, False, False, False, False, False, False, False
                    if "Rationale:" in line:
                        flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs, flagControl, flagMetre  = False, False, False, True, False, False, False, False, False, False, False
                    if "Impact:" in line:
                        flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs, flagControl, flagMetre  = False, False, False, False, True, False, False, False, False, False, False
                    # On a handful of benchmark items, the string Audit: does appear in the description. So, I'll do this one differently.
                    # In restrospect, this may have been a better way to do it anyway.
                    if re.match(r"^Audit:", line):
                    #if "Audit:" in line:
                        flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs, flagControl, flagMetre  = False, False, False, False, False, True, False, False, False, False, False
                    if "Remediation:" in line:
                        flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs, flagControl, flagMetre  = False, False, False, False, False, False, True, False, False, False, False
                    if "Default Value:" in line:
                        flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs, flagControl, flagMetre  = False, False, False, False, False, False, False, True, False, False, False
                    if "References:" in line:
                        flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs, flagControl, flagMetre  = False, False, False, False, False, False, False, False, True, False, False
                    if "CIS Controls:" in line:
                        # If we see this string, we want to make sure we're dumping out of our item. Sometimes we get here early due to no References section.
                        flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs, flagControl, flagMetre  = False, False, False, False, False, False, False, False, False, True, False
                        flagComplete = True
                        flagSkipToPageBreak = True

                    if "MITRE ATT&CK Mappings:" in line:
                        flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs, flagControl, flagMetre  = False, False, False, False, False, False, False, False, False, False, True
                       

                    if "This section is intentionally blank and exists to ensure" in line:
                        # If we see this string, we want to make sure we're dumping out of our item and resetting everything. This is a blank item.
                        flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs,flagControl, flagMetre = False, False, False, False, False, False, False, False, False, False, False
                        cis_name, cis_level, cis_desc, cis_ration, cis_impact, cis_audit, cis_remed, cis_defval, cis_refs,cis_control, cis_metre = "","","","","","","","","", "", ""
                        # flagComplete = True
                        flagStart = False
                        continue
                    if "This section contains" in line:
                        # If we see this string, we want to make sure we're dumping out of our item and resetting everything. This is a blank item.
                        flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs,flagControl, flagMetre = False, False, False, False, False, False, False, False, False, False, False
                        cis_name, cis_level, cis_desc, cis_ration, cis_impact, cis_audit, cis_remed, cis_defval, cis_refs,cis_control, cis_metre = "","","","","","","","","", "", ""
                        # flagComplete = True
                        flagStart = False
                        continue


                    if flagName:    # Here we are stitching together our entries based on the section we're in
                        cis_name = cis_name + line
                    if flagLevel:
                        cis_level = cis_level + line.replace('Profile Applicability: \n','').replace('•  ','').replace(' \n','\n')
                    if flagDesc:
                        cis_desc = cis_desc + line.replace('Description: \n','')
                    if flagRation:
                        cis_ration = cis_ration + line.replace('Rationale: \n','')
                    if flagImpact:
                        cis_impact = cis_impact + line.replace('Impact: \n','')
                    if flagAudit:
                        cis_audit = cis_audit + line.replace('Audit: \n','')
                    if flagRemed:
                        cis_remed = cis_remed + line.replace('Remediation: \n','')
                    if flagDefval:
                        cis_defval = cis_defval + line.replace('Default Value: \n','')
                    if flagRefs:
                        if "Additional Information" in line:
                            flagRefs = False
                            continue
                        line = line.replace('References: ', '').strip()
                        if re.match(r'^\d+\.\s', line):
                            cis_refs += "\n" + line
                            
                        else:
                            cis_refs += "" + line
                    if flagControl:
                        cis_control = cis_control + line.replace('CIS Controls: \n','')

                    if flagMetre:
                        cis_metre = cis_metre + line.replace('MITRE ATT&CK Mappings: \n','')

                    

        # Save to JSON
        print(f"[+] Writing JSON file '{cisjson}'...")
        with open(cisjson, 'w', encoding='utf-8') as json_file:
            json.dump(listObj, json_file, indent=4, separators=(',', ': '))

        # Convert JSON data to Excel
        print(f"[+] Creating Excel file '{cisexcel}'...")
        df_json = pd.DataFrame(listObj)  # Convert list of dicts to DataFrame
        df_json.to_excel(cisexcel, index=False)
        
        print(f"[+] Processing of '{filename}' complete!\n")

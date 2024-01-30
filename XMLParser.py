import xml.etree.ElementTree as ET
from datetime import datetime
import pandas as pd


def create_excel(data, excel_file):
    # Convert the data to a Pandas DataFrame
    df = pd.DataFrame(data)

    # Write the DataFrame to an Excel file
    df.to_excel(excel_file, index=False)


def indent(elem, level=0):
   # Add indentation
   indent_size = "  "
   i = "\n" + level * indent_size
   if len(elem):
      if not elem.text or not elem.text.strip():
         elem.text = i + indent_size
      if not elem.tail or not elem.tail.strip():
         elem.tail = i
      for elem in elem:
         indent(elem, level + 1)
      if not elem.tail or not elem.tail.strip():
         elem.tail = i
   else:
      if level and (not elem.tail or not elem.tail.strip()):
         elem.tail = i

def pretty_print_xml_elementtree(xml_string):
   # Parse the XML string
   root = ET.fromstring(xml_string)

   # Indent the XML
   indent(root)

   # Convert the XML element back to a string
   pretty_xml = ET.tostring(root, encoding="unicode")

   # Print the pretty XML
   print(pretty_xml)

def parse_xml(xml_file, voucher_tag, childs_tags, child_sub_tag, other_tag, fields, parents_fields_na, childs_fields_na, others_fields_na):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    data = []
    for voucher in root.findall(".//" + voucher_tag):  
        voucher_date = datetime.strftime(datetime.strptime(voucher.find("DATE").text, '%Y%m%d').date(), "%d-%m-%Y")
        voucher_number = voucher.find("VOUCHERNUMBER").text
        voucher_type = voucher.find("VOUCHERTYPENAME").text 
        ledger_name = ""
        if voucher.iterfind("PARTYLEDGERNAME"):
            for x in voucher.iterfind("PARTYLEDGERNAME"):
                ledger_name = x.text
                if ledger_name:
                    break  
        if not ledger_name:
            if voucher.iterfind("PARTYNAME"):
                for x in voucher.iterfind("PARTYNAME"):
                    ledger_name = x.text
                    if ledger_name:
                        break
        if not ledger_name:
            continue

        
        child_entries = []
        parent_amount = 0.0
        child_elements_entries = []
        for child_tag in childs_tags:
            for child_elements in voucher.findall(".//" + child_tag): 
                child_elements_entries.append(child_elements)
        for child_elements in child_elements_entries: 
            # child_bills = [elem.tag for elem in child_elements.iter()]
            # if len(child_bills) > 1 :
            child_ledger_name = child_elements.find("LEDGERNAME").text
            for child in child_elements.findall(".//" + child_sub_tag):
                bills = [elem.tag for elem in child.iter()]
                if len(bills) > 1:
                    child_amount = child.find("AMOUNT").text
                    ref_no = child.find("NAME").text
                    ref_type = child.find("BILLTYPE").text
                    ref_date = child.find("REFERENCEDATE").text if child.find("REFERENCEDATE") else ""
                    
                    child_data = {
                        "Date": voucher_date,
                        "Transaction Type" : "Child",
                        "Vch No.": voucher_number,
                        "Ref No": ref_no,
                        "Ref Type": ref_type,
                        "Ref Date": ref_date,
                        "Debtor": child_ledger_name,                    
                        "Ref Amount": child_amount,
                        "Amount": "NA",
                        "Particulars": child_ledger_name,
                        "Vch Type": voucher_type,
                        "Amount Verified": "NA"
                    }
                    # add amount only for child
                    parent_amount += float(child_amount)
                    child_entries.append(child_data)
            verify_amount = 0.0
            for other in child_elements.findall(".//" + other_tag):
                # print("other")
                other_bills = [elem.tag for elem in other.iter()]
                if len(other_bills) > 1:
                    other_amount = other.find("AMOUNT").text
                    other_data = {
                        "Date": voucher_date,
                        "Transaction Type" : "Other",
                        "Vch No.": voucher_number,
                        "Ref No": "NA",
                        "Ref Type": "NA",
                        "Ref Date": "NA",
                        "Debtor": child_ledger_name,                    
                        "Ref Amount": "NA",
                        "Amount": other_amount,
                        "Particulars": child_ledger_name,
                        "Vch Type": voucher_type,
                        "Amount Verified": "NA"
                    }
                    verify_amount += float(other_amount)
                    child_entries.append(other_data)

        if parent_amount == verify_amount:
            amount_verified = "Yes"
        else:
            amount_verified = "No"
        voucher_data = {
            "Date": voucher_date,
            "Transaction Type" : "Parent",           
            "Vch No.": voucher_number,
            "Ref No": "NA",
            "Ref Type": "NA",
            "Ref Date": "NA",
            "Debtor": ledger_name,
            "Ref Amount": "NA",
            "Amount": str(parent_amount),
            "Particulars": ledger_name,
            "Vch Type": voucher_type,
            "Amount Verified": amount_verified          
        }
        # Append data to the list
        data.append(voucher_data)
        data.extend(child_entries)
       
    return data

xml_file_path = 'Input.xml'
excel_file_path = 'Result.xlsx'



###############################################################################################################
# Customize these variables based on the structure of your XML file
voucher_tag = "VOUCHER"  # Tag representing each transaction

# Fields to extract
fields = ["Date", "Transaction Type", "Vch No.", "Ref No", "Ref Type", "Ref Date",            
          "Debtor", "Ref Amount", "Amount", "Particulars", "Vch Type", "Amount Verified"]

# Parents non-specific Fields
parents_fields_na = ["Ref No", "Ref Type", "Ref Date", "Ref Amount"]

# Child non-Specific Fields
childs_fields_na = ["Amount", "Amount Verified"]
childs_tags = ["ALLLEDGERENTRIES.LIST", "LEDGERENTRIES.LIST"]
child_sub_tag = "BILLALLOCATIONS.LIST"

# Others non-Specific Fields
others_fields_na = ["Ref No", "Ref Type", "Ref Date","Ref Amount", "Amount Verified"]
other_tag = "BANKALLOCATIONS.LIST"

##############################################################################################################


if __name__ == "__main__":
    parsed_data = parse_xml(xml_file_path, voucher_tag, childs_tags, child_sub_tag, other_tag, fields, parents_fields_na, childs_fields_na, others_fields_na)
    create_excel(parsed_data, excel_file_path)
    print("DONE")
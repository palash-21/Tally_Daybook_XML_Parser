from flask import Flask, request, jsonify
from XMLParser import parse_xml, create_excel, validate_xml
from datetime import datetime

app = Flask(__name__)

@app.route('/parse-xml', methods=['POST'])
def parse_xml_endpoint():
    try:
        # Receive the XML file as a POST request
        xml_content = request.files['file'].read()

        # Validate the xml
        validation, msg = validate_xml(xml_content)
        if not validation:
            return jsonify({"data": "", "error": msg})

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

        parsed_data = parse_xml(xml_content, voucher_tag, childs_tags, child_sub_tag, other_tag, fields, parents_fields_na, childs_fields_na, others_fields_na))

        # Optional save to excel file
        excel_file = 'Result.xlsx'
        excel_file += str(datetime.now())
        create_excel(parsed_data, excel_file)
        return jsonify({"data": parsed_data})
    
    except Exception as e:
        return jsonify({"error": str(e)})

if __name__ == '__main__':
    app.run(debug=True)
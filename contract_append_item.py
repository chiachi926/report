import pandas as pd
import numpy as np
import contract_grab_json

file_path = "/Users/chiachi/Documents/Master/0706/data.xls"
df = pd.read_excel(file_path)

# 顯示前幾行數據來了解結構
# print(df.head())

while True:
    contract_number = input("請輸入生效合約編號/市調合約編號: ")

    if contract_number == "exit":
        break

    # 根據生效合約編號篩選數據
    # contract_data = df[df["生效合約編號"] == contract_number]  # 修改這裡以匹配實際欄位名稱
    contract_data = df[(df["生效合約編號"] == contract_number) | (df["市調合約編號"] == contract_number)]
    print(contract_data)
    index = int(input("請選擇編號："))

    results=[]

    if contract_data.empty:
        print("找不到相應的生效合約編號")
    else:
        # 部門
        department_conditions = {
            "全企業": ["通用性/總處督導公司合約", "合約期間內調整合約（排除運輸合約）"],
            "單一公司": ["NTD 200萬元以上長庚醫院合約", "NTD 200萬元以上單一督導公司合約", "NTD 20萬元以上生醫體系合約"]
        }

        def determine_value(value, conditions):
            for key, condition_values in conditions.items():
                if value in condition_values:
                    return key
            return None
        
        # 新舊約
        def determine_contract_type(value):
            return "新合約" if "*" in value else "舊合約" 
        
        # 材料項次
        def determine_material_type(value):
            return "單項材料" if value==1 else "多項材料"
        
        # 詢價
        def determine_inquiry_type(value):
            return "獨家報價" if value==1 else "多家報價"
        
        # 議價
        def determine_negotiation_type(value, contract_type, material_type, inquiry_type):
            if value=="      ": 
                return "無上次訂約" #0
            # elif determine_contract_type(value)=="舊合約" and determine_material_type(value)=="多項材料" and determine_inquiry_type(value)=="獨家報價":
            elif contract_type == "舊合約" and material_type == "多項材料" and inquiry_type == "獨家報價":
                return "訂約一家" #2
            else:
                return "有上次訂約" #1

        # 金額比較
        def determine_price_comparison_type(value):
            if value>0.00:
                return "較前一般採購價高"
            elif value==0.00:
                return "同前一般採購價"
            elif value<0.00:
                return "較前一般採購價低"
            else:
                return "無前一般採購價"


        department = np.vectorize(determine_value)(contract_data["呈核類型名稱"].values[index-1], department_conditions)
        if department=="全企業":
            results.append("0")
        elif department=="單一公司":
            results.append("1")

        contract_type = np.vectorize(determine_contract_type)(contract_data["市調合約編號"].values[index-1])
        if contract_type=="新合約":
            results.append("0")
        elif contract_type=="舊合約":
            results.append("1")
        
        material = np.vectorize(determine_material_type)(contract_data["規格數量"].values[index-1])
        if material=="單項材料":
            results.append("0")
        elif material=="多項材料":
            results.append("1")
        
        inquiry = np.vectorize(determine_inquiry_type)(contract_data["報價廠商數"].values[index-1])
        if inquiry=="獨家報價":
            results.append("0")
        elif inquiry=="多家報價":
            results.append("1")

        negotiation = np.vectorize(determine_negotiation_type)(contract_data["前次訂約廠商簡稱"].values[index-1], contract_type, material, inquiry)
        if negotiation=="無上次訂約":
            results.append("0")
        elif negotiation=="有上次訂約":
            results.append("1")
        else: #訂約一家
            results.append("2")
        
        price_comparison = np.vectorize(determine_price_comparison_type)(contract_data["議價調幅"].values[index-1])
        if price_comparison=="較前一般採購價高":
            results.append("0")
        elif price_comparison=="同前一般採購價":
            results.append("1")
        elif price_comparison=="較前一般採購價低":
            results.append("2")
        elif price_comparison=="無前一般採購價":
            results.append("3")

        print(f"使用部門: {department}")
        print(f"新舊約: {contract_type}")
        print(f"材料項次: {material}")
        print(f"詢價: {inquiry}")
        print(f"議價: {negotiation}")
        print(f"金額比較: {price_comparison}\n")
        #print(results)

        item=0
        if results==list("000000"):
            item=1
        elif results==list("000001"):
            item=2
        elif results==list("000002"):
            item=3
        elif results==list("000003"):
            item=4
        elif results==list("000100"):
            item=5
        elif results==list("000101"):
            item=6
        elif results==list("000102"):
            item=7
        elif results==list("000103"):
            item=8
        elif results==list("001000"):
            item=9
        elif results==list("001001"):
            item=10
        elif results==list("001002"):
            item=11
        elif results==list("001003"):
            item=12
        elif results==list("001100"):
            item=13
        elif results==list("001101"):
            item=14
        elif results==list("001102"):
            item=15
        elif results==list("001103"):
            item=16
        elif results==list("010010"):
            item=17
        elif results==list("010011"):
            item=18
        elif results==list("010012"):
            item=19
        elif results==list("010110"):
            item=20
        elif results==list("010111"):
            item=21
        elif results==list("010112"):
            item=22
        elif results==list("011020"):
            item=23
        elif results==list("011021"):
            item=24
        elif results==list("011022"):
            item=25
        elif results==list("011110"):
            item=26
        elif results==list("011111"):
            item=27
        elif results==list("011112"):
            item=28
        elif results==list("100000"):
            item=29
        elif results==list("100001"):
            item=30
        elif results==list("100002"):
            item=31
        elif results==list("100003"):
            item=32
        elif results==list("100100"):
            item=33
        elif results==list("100101"):
            item=34
        elif results==list("100102"):
            item=35
        elif results==list("100103"):
            item=36
        elif results==list("101000"):
            item=37
        elif results==list("101001"):
            item=38
        elif results==list("101002"):
            item=39
        elif results==list("101003"):
            item=40
        elif results==list("101100"):
            item=41
        elif results==list("101101"):
            item=42
        elif results==list("101102"):
            item=43
        elif results==list("101103"):
            item=44
        elif results==list("110010"):
            item=45
        elif results==list("110011"):
            item=46
        elif results==list("110012"):
            item=47
        elif results==list("110110"):
            item=48
        elif results==list("110111"):
            item=49
        elif results==list("110112"):
            item=50
        elif results==list("111020"):
            item=51
        elif results==list("111021"):
            item=52
        elif results==list("111022"):
            item=53
        elif results==list("111110"):
            item=54
        elif results==list("111111"):
            item=55
        elif results==list("111112"):
            item=56

        print(f"項次={item}")


    data = contract_grab_json.grab_data_from_excel(file_path, contract_number, index)
    if data:
        report1 = contract_grab_json.generate_audit_report(item, data)
        print(f"{report1}\n")



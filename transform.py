import pandas as pd
import requests
import time
from tqdm import tqdm  # 可选，用于显示进度条


def get_smiles_from_cas(cas):
    """
    通过CAS编号查询SMILES字符串
    """
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    }

    try:
        # 第一步：通过CAS获取CID
        url_cid = f'https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/{cas}/cids/JSON'
        response = requests.get(url_cid, headers=headers)

        if response.status_code != 200:
            return None

        cid = response.json()['IdentifierList']['CID'][0]

        # 添加延迟以避免频繁请求
        time.sleep(0.2)

        # 第二步：通过CID获取SMILES
        url_property = f'https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/cid/{cid}/property/SMILES/JSON'
        response = requests.get(url_property, headers=headers)

        if response.status_code == 200:
            return response.json()['PropertyTable']['Properties'][0]['SMILES']
        return None

    except Exception as e:
        print(f"Error processing CAS {cas}: {str(e)}")
        return None


# 读取Excel文件
input_file = "test_cas.xlsx"  # 输入文件名
output_file = "compounds_with_smiles.xlsx"  # 输出文件名

df = pd.read_excel(input_file)

# 确保CAS列存在
if 'CAS' not in df.columns:
    raise ValueError("Excel文件中必须包含'CAS'列")

# 添加新列（如果不存在）
if 'SMILES' not in df.columns:
    df['SMILES'] = None

# 遍历每一行获取SMILES
for index, row in tqdm(df.iterrows(), total=len(df)):
    if pd.isnull(row['SMILES']) or row['SMILES'] == '':
        cas = str(row['CAS']).strip()
        if cas != 'nan':  # 跳过空值
            df.at[index, 'SMILES'] = get_smiles_from_cas(cas)

# 保存结果到新文件
df.to_excel(output_file, index=False)
print(f"处理完成！结果已保存到 {output_file}")
import pandas as pd
import requests
import time
from tqdm import tqdm
from urllib.parse import quote
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry

# ================== 配置参数 ==================
INPUT_FILE = "298.15k下的液态热容.xlsx"  # 输入文件名
OUTPUT_FILE = "yetai_CP.xlsx"  # 输出文件名
SHEET_NAME = "2040"  # 工作表名称
CAS_COLUMN = "CAS"  # CAS号列名
SMILES_COLUMN = "SMILES"  # 结果列名
MAX_RETRIES = 3  # 最大重试次数
RETRY_DELAY = 3  # 重试基础等待时间(秒)
REQUEST_TIMEOUT = 15  # 请求超时时间(秒)
REQUEST_INTERVAL = 1.2  # 请求间隔时间(秒)

# ================== 高级请求配置 ==================
retry_strategy = Retry(
    total=MAX_RETRIES,
    backoff_factor=RETRY_DELAY,
    status_forcelist=[429, 500, 502, 503, 504],
    allowed_methods=["GET"]
)


def create_session():
    """创建带重试机制的会话"""
    session = requests.Session()
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
        'Accept': 'application/json'
    })
    return session


def clean_cas(cas):
    """清洗CAS号数据"""
    if pd.isna(cas):
        return None
    cas = str(cas).strip()
    # 过滤无效CAS格式（简单校验）
    if len(cas) < 5 or cas.count("-") != 2:
        return None
    return cas


def get_smiles(session, cas):
    """通过CAS号查询SMILES"""
    try:
        # 第一步：获取CID
        encoded_cas = quote(cas)
        cid_url = f"https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/{encoded_cas}/cids/JSON"

        response = session.get(cid_url, timeout=REQUEST_TIMEOUT)
        if not response.ok:
            return None

        cid = response.json()['IdentifierList']['CID'][0]

        # 第二步：获取SMILES
        smiles_url = f"https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/cid/{cid}/property/SMILES/JSON"
        response = session.get(smiles_url, timeout=REQUEST_TIMEOUT)

        return response.json()['PropertyTable']['Properties'][0]['SMILES'] if response.ok else None

    except Exception as e:
        # print(f"Error processing {cas}: {str(e)}")  # 调试时启用
        return None


def main():
    # 读取数据
    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)

    # 初始化列
    if SMILES_COLUMN not in df.columns:
        df[SMILES_COLUMN] = None

    # 创建会话
    session = create_session()

    # 进度条设置
    pbar = tqdm(total=len(df), desc="Processing CAS numbers")

    # 批量处理
    for idx, row in df.iterrows():
        current_cas = clean_cas(row[CAS_COLUMN])

        # 跳过无效CAS号
        if not current_cas or (pd.notna(row[SMILES_COLUMN]) and row[SMILES_COLUMN]):
            pbar.update(1)
            continue

        # 执行查询
        smiles = get_smiles(session, current_cas)
        df.at[idx, SMILES_COLUMN] = smiles or ""  # 避免NaN

        # 进度更新
        pbar.update(1)
        time.sleep(REQUEST_INTERVAL)  # 遵守速率限制

    pbar.close()

    # 保存结果
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n处理完成！结果已保存到 {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
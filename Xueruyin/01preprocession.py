#载入模块
import warnings, re
import pandas as pd, numpy as np

# 忽略所有警告
warnings.filterwarnings("ignore")

# =======================
# 1. 读取数据
# =======================
# 设置要读取的列范围
cols = ["E:T", "W:Y", "AA:AL", "AM:AV", "AZ:BC", "BE", "BI", "BK:BQ", "CC:CD", "CH:CI"]
usecols = ",".join(cols)

# 读取 Excel 数据并设置表头行（从第二行开始）
df = pd.read_excel(r"00安全性分析数据-乙肝.xlsx",usecols=usecols,header=1)

# =======================
# 2. 更新列名
# =======================
# 使用映射字典将中文列名转换为英文列名
new_cols = {
    '民族': 'Ethnicity', 
    '年龄': 'Age', 
    '婚姻状况': 'Marital status', 
    '丈夫年龄': 'Husband age', 
    '文化程度': "Education", 
    '职业': 'Occupation', 
    '户籍类型': 'Local', 
    '是否广东户籍': 'Guangdong', 
    '末次月经时间':'Last menstrual time',
    '孕次': 'Gravida', 
    '产次': 'Parity',
    '初检孕周': 'First check pregnancy week', 
    '分娩孕周': 'Delivery pregnancy week', 
    '产妇数': 'Number of deliveries', 
    '胎数': 'Number of fetuses', 
    '分娩方式': 'Delivery method', 
    '妊娠合并症': 'Gestational complications', 
    '产检总次数': 'Total antenatal visits', 
    '分娩医院': 'Delivery hospital', 
    'HBsAg': 'HBsAg at first', 
    '抗-HBs': 'Anti-HBs at first', 
    'HBeAg': 'HBeAg at first', 
    '抗-HBe': 'Anti-HBe at first', 
    '抗-HBc': 'Anti-HBc at first', 
    '首次检测孕周': 'First test pregnancy week', 
    'HBsAg.1': 'HBsAg at delivery', 
    '抗-HBs.1': 'Anti-HBs at delivery',
    'HBeAg.1': 'HBeAg at delivery', 
    '抗-HBe.1': 'Anti-HBe at delivery', 
    '抗-HBc.1': 'Anti-HBc at delivery', 
    '孕晚期或分娩前后检测孕周': 'Late pregnancy test week', 
    '有无DNA检测': 'Whether DNA test was conducted',
    '首次DNA检测结果': 'First DNA test result',
    '首次DNA检测时间': 'Time of first DNA test',
    '首次DNA检测孕周': 'Gestation week of first DNA test',
    '首检DNA是否高病载': 'High viral load at first test',
    '最后一次DNA检测': 'Whether Last DNA test was conducted',
    '最后一次DNA检测结果': 'Last DNA test result',
    '最后一次DNA检测时间': 'Time of last DNA test',
    '最后一次DNA检测孕周': 'Gestation week of last DNA test',
    '最后一次DNA是否高病载': 'High viral load in last test',
    'ALT': 'ALT', 
    'AST': 'AST', 
    'TBIL': 'TBIL', 
    'Unnamed: 54': 'Drug', 
    'Unnamed: 56': 'Drug starting week', 
    '性别': 'Baby gender', 
    '转归': 'Baby outcome', 
    '体重': 'Baby weight',
    '身长': 'Baby height', 
    '头围（cm）': 'Baby head circumference (cm)', 
    'Apgar1': 'Apgar1', 
    'Apgar5': 'Apgar5', 
    '有无出生缺陷': 'Birth defects', 
    'HBsAg.2': 'Baby HBsAg after birth', 
    'HBsAb': 'Baby HBsAb after birth',
    '孕妇有无不良事件': 'Adverse events', 
    '不良事件内容': 'Adverse event details'
}

# 更新 DataFrame 的列名并处理数据中的 "/" 和 NaN 值
df.rename(columns=new_cols, inplace=True)
df.replace({"nan": np.nan, "/": np.nan}, inplace=True)

# =======================
# 3. 转换列数据
# =======================
#民族
df['Ethnicity'] = df['Ethnicity'].apply(lambda x: 'Han' if '汉' in str(x) else 'Others')

#婚姻状况
marital_status_mapping = {
    '再婚': 'Married',
    '初婚': 'Married',
    '未婚': 'Unmarried',
    '同居': 'Unmarried',
    '离婚': 'Divorced',
    '丧偶': 'Widowed'
}
df['Marital status']=df['Marital status'].replace(marital_status_mapping)

#文化程度
education_mapping = {
    '文盲': 'Illiterate',
    '小学': 'Primary school',
    '初中': 'Secondary school',
    '中专': 'Secondary school',
    '高中': 'Secondary school',
    '技工': 'Secondary school',
    '大专': 'Tertiary',
    '大学': 'Tertiary',
    '本科': 'Tertiary',
    '硕士': 'Postgraduate',
    '博士': 'Postgraduate'
}
df['Education'] = df['Education'].replace(education_mapping)

# 职业
df['Occupation'] = df['Occupation'].replace({
                                            # Employed
                                            '干部': 'Employed',
                                            '干部职员': 'Employed',
                                            '职员': 'Employed',
                                            '其他：职员': 'Employed',
                                            '商业服务': 'Employed',
                                            '商业': 'Employed',
                                            '工人': 'Employed',
                                            '工人、农民工': 'Employed',
                                            '农民': 'Employed',
                                            '渔民': 'Employed',
                                            '医务人员': 'Employed',
                                            '教师': 'Employed',
                                            '保育员及保姆': 'Employed',
                                            '餐饮食品业': 'Employed',
                                            # Self-employed
                                            '个体': 'Self-employed',
                                            '其他：个体': 'Self-employed',
                                            # Unemployed / Retired
                                            '家务': 'Unemployed/Retired',
                                            '家务及待业': 'Unemployed/Retired',
                                            '离退人员': 'Unemployed/Retired',
                                            # Students
                                            '学生': 'Students',
                                            # Others
                                            '其他': 'Unemployed/Retired',
                                            '不详':np.nan})

#户籍
df['Local'] = df['Local'].replace({
    '深圳户籍': 'Local',
    '流动': 'Non-local',
    '暂住': 'Non-local'
})

#广东省内
df['Guangdong'] = df['Guangdong'].replace({
    '是': 'Yes',
    '否': 'No'
})

#末次月经时间
df['Last menstrual time'] = pd.to_datetime(df['Last menstrual time'], errors='coerce').dt.strftime('%Y-%m-%d')

#孕次、产次、初检孕周、分娩孕周、产妇数、胎数
columns_to_convert=[
    'Gravida', 
    'Parity', 
    'First check pregnancy week', 
    'Delivery pregnancy week', 
    'Number of deliveries', 
    'Number of fetuses'
]
for col in columns_to_convert:
    valid_rows = df[[col]].replace([np.inf, -np.inf], np.nan).notna().all(axis=1)
    df.loc[valid_rows, col] = df.loc[valid_rows, col].astype(int)
    
# 转换Parity列
bins = [-np.inf, 0, 3, np.inf]
labels = ['0', '1-3', '>3']
df['Parity'] = pd.cut(df['Parity'], bins=bins, labels=labels)

#分娩方式
mapping = {
    '产钳': 'Forceps delivery',
    '胎头吸引': 'Vacuum extraction',
    '剖宫产': 'Cesarean section',
    '顺产': 'Natural delivery',
    '臀位': 'Breech'
}
df['Delivery method'] = df['Delivery method'].map(mapping).astype('category')

#妊娠合并症
mapping = {
    # 直接映射
    '胎盘早剥':'Placental abruption', '羊水栓塞':'Amniotic fluid embolism',
    '产后出血≥500ml':'Postpartum hemorrhage', '胎盘植入':'Placenta accreta',
    '围产儿死亡':'Perinatal death', '产褥感染':'Puerperal infection',
    '早产':'Preterm birth', '胎膜早破':'Premature rupture of membranes',
    '先兆子痫':'Threatened eclampsia', '出生缺陷':'Birth defect',
    '胎盘滞留':'Placenta retention', '前置胎盘':'Placenta previa',
    '妊娠高血压疾病':'Gestational hypertension', '慢性高血压':'Gestational hypertension',
    '高血压病':'Gestational hypertension', '妊娠糖尿病':'Gestational diabetes',
    '糖尿病':'Gestational diabetes', '先兆子宫破裂':'Uterine rupture',
    '子宫破裂':'Uterine rupture','子痫':'Eclampsia','子痫前期轻度':'Eclampsia', '子痫前期重度':'Eclampsia'
}
exclude = {'肝炎','肝病','其他','心脏病'}
df['Gestational complications'] = (df['Gestational complications'].fillna('').str.split(';').apply(lambda items: ';'.join(dict.fromkeys(mapping.get(it.strip(), it.strip()) for it in items if (s:=it.strip()) and s not in exclude))).replace('', pd.NA))
df = (df.drop('Gestational complications', axis=1).join(df['Gestational complications'].fillna('').str.get_dummies(sep=';')))

#分娩医院
hospital_mapping = {"南方医科大学深圳医院": "General Hospital",
                    "深圳万丰医院": "General Hospital",
                    "深圳同仁妇产医院": "Obstetrics & Gynecology Hospital",
                    "深圳复亚医院": "General Hospital",
                    "深圳天伦医院": "Specialty Hospital",
                    "深圳宝生妇产医院": "Obstetrics & Gynecology Hospital",
                    "深圳市中西医结合医院": "TCM and Western Medicine Hospital",
                    "深圳市宝安区中医院": "TCM and Western Medicine Hospital",
                    "深圳市宝安区中心医院": "District General Hospital",
                    "深圳市宝安区人民医院": "District General Hospital",
                    "深圳市宝安区妇幼保健院": "Obstetrics & Gynecology Hospital",
                    "深圳市宝安区松岗人民医院": "District General Hospital",
                    "深圳市宝安区石岩人民医院": "District General Hospital",
                    "深圳市宝安区福永人民医院": "District General Hospital",
                    "深圳广生医院": "Specialty Hospital",
                    "深圳恒生医院": "General Hospital",
                    "深圳永福医院": "Specialty Hospital",
                    "深圳燕罗塘医院": "Specialty Hospital"}
df['Delivery hospital'] = df['Delivery hospital'].map(hospital_mapping).fillna('Other').astype('category')

# 检验定性数据
columns_to_check = [
    'HBsAg at first', 'Anti-HBs at first', 'HBeAg at first', 'Anti-HBe at first', 'Anti-HBc at first',
    'HBsAg at delivery', 'Anti-HBs at delivery', 'HBeAg at delivery', 'Anti-HBe at delivery', 'Anti-HBc at delivery',
    'Baby HBsAg after birth', 'Baby HBsAb after birth'
]
for col in columns_to_check:
    df[col] = df[col].apply(lambda x: 'Negative' if pd.notna(x) and '阴' in str(x) else ('Positive' if pd.notna(x) else x))
    
# 定义大小三阳函数
def classify_hbv_status(hbsag, hbeag, hbcab):
    if pd.isna(hbsag) or pd.isna(hbeag) or pd.isna(hbcab):
        return pd.NA
    if hbsag == 'Positive' and hbcab == 'Positive':
        if hbeag == 'Positive':
            return 'Major triad positive'
        elif hbeag == 'Negative':
            return 'Minor triad positive'
    return 'None'
#转换大小三阳
df['HBV status at first'] = df.apply(lambda row: classify_hbv_status(row['HBsAg at first'],row['HBeAg at first'],row['Anti-HBc at first']),axis=1)
df['HBV status at delivery'] = df.apply(lambda row: classify_hbv_status(row['HBsAg at delivery'],row['HBeAg at delivery'],row['Anti-HBc at delivery']),axis=1)

# 首次检测孕周/孕晚期或分娩前后检测孕周/首次DNA检测孕周/最后一次DNA检测孕周
for col in ['Late pregnancy test week', 'First test pregnancy week', 'Gestation week of first DNA test', 'Gestation week of last DNA test']:
    df[col] = df[col].astype(str).str.replace('产后', '43').str.replace('不详', '').str.replace('孕前', '0').replace('', np.nan)

# 有无DNA检测
df['Whether DNA test was conducted'] = df['Whether DNA test was conducted'].fillna("No").replace({'无':"No", '有':"Yes"})
df['Whether Last DNA test was conducted'] = df['Whether Last DNA test was conducted'].fillna("No").replace({'无':"No", '有':"Yes"})

# 首次DNA检测结果列/最后一次DNA检测结果列
def log_transform_dna_column(col):
    extracted = df[col].astype(str).str.extract(r'(\d+\.?\d*)')[0]
    df[col] = pd.to_numeric(extracted, errors='coerce')
    df[col] = df[col].apply(lambda x: np.log10(x) if pd.notna(x) and x > 0 else x)
log_transform_dna_column('First DNA test result')
log_transform_dna_column('Last DNA test result')
# 确保两列都不为空时才做减法，否则结果为 NaN
mask = df[['First DNA test result', 'Last DNA test result']].notna().all(axis=1)
df['Viral decrease'] = np.nan  # 初始化新列为 NaN
df.loc[mask, 'Viral decrease'] = df.loc[mask, 'First DNA test result'] - df.loc[mask, 'Last DNA test result']

# 首次DNA检测时间/最后一次DNA检测时间
df['Time of first DNA test'] = df['Time of first DNA test'].replace('不详', pd.NaT)
df['Time of last DNA test'] = df['Time of last DNA test'].replace('不详', pd.NaT)
def excel_date_to_datetime(x):
    try:
        x_num = float(x)
        return pd.to_datetime('1899-12-30') + pd.Timedelta(days=x_num)
    except:
        return pd.NaT 
df['Time of first DNA test'] = df['Time of first DNA test'].apply(excel_date_to_datetime)
df['Time of last DNA test'] = df['Time of last DNA test'].apply(excel_date_to_datetime)
df['Time of first DNA test'] = df['Time of first DNA test'].dt.strftime('%Y-%m-%d')
df['Time of last DNA test'] = df['Time of last DNA test'].dt.strftime('%Y-%m-%d')

# 高病毒载量列
virus_cols={'是':'Yes','否':'No','不详':np.nan}
df['High viral load at first test']=df['High viral load at first test'].replace(virus_cols)
df['High viral load in last test']=df['High viral load in last test'].replace(virus_cols)

# ALT
df['ALT'] = df['ALT'].replace('<', '', regex=True)
df['ALT'] = pd.to_numeric(df['ALT'], errors='coerce')  # 将无法转换的值转为 NaN

# AST
df['AST'] = df['AST'].replace('不详', np.nan)
df['AST'] = df['AST'].str.extract(r'(\d+\.?\d*)')  # 提取其中的数字部分
df['AST'] = pd.to_numeric(df['AST'], errors='coerce')  # 转换为数值

# TBIL
df['TBIL'] = df['TBIL'].str.extract(r'(\d+\.?\d*)')  # 提取其中的数字部分
df['TBIL'] = pd.to_numeric(df['TBIL'], errors='coerce')  # 转换为数值

# 用药情况
df['Drug'] = df['Drug'].replace({'否': 'No', '无 ': 'No', '是': 'Yes', '有': 'Yes'})

# 用药开始时间
df['Drug starting week'] = df['Drug starting week'].astype(str)
df['Drug starting week'] = df.apply(
    lambda row:
        0 if any(x in row['Drug starting week'] for x in ['前', '开始', '一直', '长期', '整个']) else
        np.nan if any(x in row['Drug starting week'] for x in ['不详', '产时', '未用药']) else
        43 if '产后' in row['Drug starting week'] else
        4 if '详见备注' in row['Drug starting week'] else
        float(re.search(r'(\d+)\+(\d+)', row['Drug starting week']).group(1)) +
        float(re.search(r'(\d+)\+(\d+)', row['Drug starting week']).group(2)) / 7
        if re.match(r'^\d+\+\d+$', row['Drug starting week']) else
        float(re.search(r'(\d+)\+(\d+)', row['Drug starting week']).group(1)) +
        float(re.search(r'(\d+)\+(\d+)', row['Drug starting week']).group(2)) / 7
        if re.match(r'^\d+\+\d+周$', row['Drug starting week']) else
        float(re.search(r'\d+', row['Drug starting week']).group(0))
        if re.match(r'^\d+W$', row['Drug starting week']) else
        float(re.search(r'\d+', row['Drug starting week']).group(0))
        if re.match(r'^\d+周$', row['Drug starting week']) else
        (pd.to_datetime(row['Drug starting week'], errors='coerce') - pd.to_datetime(row['Last menstrual time'], errors='coerce')).days / 7
        if re.match(r'^\d{4}-\d{2}-\d{2}$', row['Drug starting week']) else
        (pd.to_datetime(row['Drug starting week'], errors='coerce') - pd.to_datetime(row['Last menstrual time'], errors='coerce')).days / 7
        if re.match(r'^\d{1,4}[-/]\d{1,2}[-/]\d{1,2}$', row['Drug starting week']) else
        float(row['Drug starting week']) if re.match(r'^\d+(\.\d+)?$', row['Drug starting week']) and float(row['Drug starting week']) < 50 else
        np.nan if re.match(r'^\d+(\.\d+)?$', row['Drug starting week']) and 50 < float(row['Drug starting week']) < 1000 else
        (pd.to_datetime(row['Drug starting week'], errors='coerce') - pd.to_datetime(row['Last menstrual time'], errors='coerce')).days / 7,
    axis=1)
bins = [0, 12, 24, float('inf')]
labels = ['Early', 'Middle', 'Late']
df['Drug start stage'] = pd.cut(df['Drug starting week'], bins=bins, labels=labels, right=True)

# 孩子性别
df['Baby gender'] = df['Baby gender'].replace("未说明的性别", np.nan)

# 孩子转归
df['Baby outcome'] = df['Baby outcome'].replace({"活产": "Live birth","死胎": "Stillbirth"})
df['Stillbirth'] = np.where(df['Baby outcome'] == 'Stillbirth', 1, 0)

# 孩子体重
more_precise_bins = [-float('inf'), 1000, 1500, 2500, 4000, float('inf')]
more_precise_labels = ['ELBW', 'VLBW', 'LBW', 'Normal', 'LGA']
df['Baby weight Category'] = pd.cut(df['Baby weight'], bins=more_precise_bins, labels=more_precise_labels, right=False)
df['LBW'] = np.where((df['Baby weight Category'] == 'LBW') | (df['Baby weight Category'] == 'VLBW') | (df['Baby weight Category'] == 'ELBW'), 1, 0)
df['LGA'] = np.where((df['Baby weight Category'] == 'LGA') | (df['Baby weight Category'] == 'Normal'), 1, 0)

# 头围
df['Baby head circumference (cm)'] = pd.to_numeric(df['Baby head circumference (cm)'], errors='coerce')
df['Baby head circumference (cm)'] = df['Baby head circumference (cm)'].where(df['Baby head circumference (cm)'] >= 10, np.nan)

# 出生缺陷
defect_mapping = {'围产儿死亡':'Perinatal death', '出生缺陷':'Birth defect'}
df['Adverse event details'] = (df['Adverse event details'].fillna('').str.split(';').apply(lambda items: ';'.join(dict.fromkeys(defect_mapping.get(it.strip(), it.strip()) for it in items if (s:=it.strip())))).replace('', pd.NA))
df = (df.drop('Adverse event details', axis=1).join(df['Adverse event details'].fillna('').str.get_dummies(sep=';')))
df['Birth defect'] = np.where((df['Birth defects'] == 1) | (df['Birth defect'] == 1), 1, 0)
df.drop(['Birth defects','Adverse events'], axis=1, inplace=True)

# 不良事件
maternal_cols = ['Amniotic fluid embolism', 'Eclampsia', 'Placenta accreta',
    'Placenta previa', 'Placenta retention', 'Placental abruption',
    'Postpartum hemorrhage', 'Premature rupture of membranes',
    'Preterm birth', 'Puerperal infection', 'Threatened eclampsia',
    'Uterine rupture']
baby_cols = ['LBW','Birth defect', 'Perinatal death', 'Stillbirth']
df['Maternal adverse outcome'] = (df[maternal_cols].sum(axis=1) > 0).astype(int)
df['Baby adverse outcome'] = (df[baby_cols].sum(axis=1) > 0).astype(int)

# 排序
outcome_cols = maternal_cols + baby_cols
other_cols = [col for col in df.columns if col not in outcome_cols]
df = df[other_cols + outcome_cols]
highlight_cols = ['Maternal adverse outcome', 'Baby adverse outcome']

#最后发现某些汉字的替换
df.replace({'有':"Yes", "无":"No","男性":"Male","女性":"Female","nan":np.nan}, inplace=True)
df.drop(columns=df.columns[df.nunique(dropna=True) == 1], inplace=True)
        
# 设置样式函数
def style_dataframe(s):
    return s.style\
        .applymap(lambda v: 'background-color: yellow', subset=highlight_cols)\
        .applymap(lambda v: 'color: red', subset=outcome_cols)
        
# 保存为带样式的 Excel 文件
styled = style_dataframe(df)
styled.to_excel(r"01data_encoding.xlsx", index=False, engine='openpyxl')
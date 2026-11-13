# -*- coding: utf-8 -*-
"""
suppliers_streamlit.py
=======================
供應商規則的統一入口。實際的 detect_xxx()/parse_xxx() 依國家拆在：
  - suppliers_indonesia_streamlit.py：印尼供應商 (ID01~)
  - suppliers_vietnam_streamlit.py：越南供應商 (VN01~VN20)
  - suppliers_common_streamlit.py：跨國共用小工具

日後新增國家 (柬埔寨/孟加拉/緬甸/中國) 時：
  1. 建一支新檔案 suppliers_<country>_streamlit.py (仿照印尼/越南的寫法，
     最後一樣用 register_supplier() 註冊)。
  2. 在這支檔案最上面加一行
       import suppliers_<country>_streamlit  # noqa: F401
  3. 同步到本機時，release_local.py 的 CORE/RENAMES 要各加一筆
     (仿照 suppliers_indonesia_streamlit.py 那一筆的寫法)。

其他程式 (app.py / verify_app.py / run_regression.py / fallback_streamlit.py)
只需要 `import suppliers_streamlit`，不用知道底下怎麼拆檔案。
"""

from extract_utils_streamlit import SUPPLIER_REGISTRY, set_supplier_number

import suppliers_indonesia_streamlit  # noqa: F401  (觸發印尼供應商 register_supplier())
import suppliers_vietnam_streamlit  # noqa: F401  (觸發越南供應商 register_supplier())

# ===========================================================================
# 供應商編號 (依進度管控的 Top 供應商清單)
# ===========================================================================
# 印尼：Desktop\印尼進度管控 的「Indonesia Supplier Top 15, 2025-2026.xlsx」
# 與「印尼進度管控表」占比工作表 (2025-2026 請款金額排名 + 廠商代號)。
#   ID01~ID15 = 印尼 TOP15；TOP15 以外沿用同一份排名 (例如 FEMARIA 第 18 名
#   = ID18)。ID08 ORIENTAL POWER LOGISTICS 的 key 是 OPL (2026-10 建立)。
# 越南：Desktop\越南進度管控 的「Vietnam Supplier Top 20.xlsx」(01~20，資料夾
#   名稱「編號-廠商代號 公司名」)，建好規則後一樣在這裡登記 VN01~VN20。
# 編號只是對照資料 (顯示在網頁清單、轉檔畫面與轉檔總表)，不影響辨識順序。

SUPPLIER_NUMBERS = {
    # 供應商 key: (編號, 廠商代號, 進度管控表上的公司名稱)
    "YJE_TATA": ("ID01", "100019748", "PT. TATA HARMONI SARANATAMA"),
    "MAERSK": ("ID02", "100019226", "PT. MAERSK LOGISTICS INDONESIA"),
    "HANDAL": ("ID03", "100019606", "PT.HANDAL MANDIRI TRANSINDO"),
    "DHL": ("ID04", "100019563", "PT. BIROTIKA SEMESTA"),
    "KUEHNE_NAGEL": ("ID05", "100019143", "PT. KUEHNE NAGEL INDONESIA"),
    "DSV": ("ID06", "100023151", "PT. DSV TRANSPORT INDONESIA"),
    "INDOPROSTIME_TRANSPORT": ("ID07", "100019300", "PT. INDOPROSTIME TRANSPORT"),
    "OPL": ("ID08", "100004377", "ORIENTAL POWER LOGISTICS"),
    "OPL_USD": ("ID08-OPL(USD)", "ID08-OPL(USD)", "OPL曜捷"),
    "DWIHARTA": ("ID09", "100019242", "PT. DWIHARTA LOGISTINDO"),
    "TRANS_DAYA_PRIMA": ("ID10", "100019488", "PT. TRANS DAYA PRIMA"),
    "GALUH": ("ID11", "100019600", "PT.GALUH INTI BAHARI"),
    "ECOMET": ("ID12", "100019489", "PT. TRANS ECOMET GLOBAL"),
    "PUTRA_SEJAHTERA_SENTOSA": ("ID13", "100030053", "PT. PUTRA SEJAHTERA SENTOSA"),
    "APL": ("ID14", "100019535", "PT. APL LOGISTICS"),
    "PANCARAN": ("ID15", "100019386", "PT. PANCARAN SRIKANDI LOGISTIK"),
    "INDOPROSTIME": ("ID16", "100030991", "PT INDO PROSTIME EXPRESS"),
    "FEMARIA": ("ID18", "100019597", "PT. FEMARIA BUANA CARGO"),
    "EXPRESS_MAXIMUM": ("ID21", "100019046", "PT EXPRESS MAXIMUM"),
    "PAN_EKSPRES": ("ID28", "100030970", "PT PAN EKSPRES INTERNATIONAL"),
    "HYPER_MEGA": ("ID31", "100019295", "PT. HYPER MEGA SHIPPING"),
    "EVERGREEN_LOGISTICS": ("ID38", "100019246", "PT. EVERGREEN LOGISTICS INDONESIA"),
    "RONG_A_CHAU": ("VN03", "100008978", "CÔNG TY TNHH GIAO NHẬN THƯƠNG MẠI RỒNG Á CHÂU"),
    "CONG_THANH": ("VN05", "100008048", "CTY TNHHDV GN VT VA TM CONG THANH"),
    "TRANS_WAGON": ("VN01", "100009281", "CÔNG TY TNHH LIÊN VẬN QUỐC TẾ"),
    "DHL_EXPRESS_VN": ("VN04", "100008681", "CÔNG TY TNHH CHUYỂN PHÁT NHANH DHL-VNPT"),
    "MAERSK_VN": ("VN07", "100008773", "CÔNG TY TNHH MAERSK LOGISTICS & SERVICES VIỆT NAM"),
    "KN_VN": ("VN10", "100009241", "CÔNG TY TNHH KUEHNE+NAGEL"),
    "APL_VN": ("VN11", "100008592", "CÔNG TY TNHH APL LOGISTICS VIỆT NAM"),
    "VINH_HUNG_LONG": ("VN02", "100008024", "CÔNG TY TNHH VĨNH HƯNG LONG"),
    "BINH_TAN": ("VN06", "100012345", "HỢP TÁC XÃ VẬN TẢI DỊCH VỤ DU LỊCH BÌNH TÂN"),
    "NGUYEN_HUONG": ("VN12", "100009955", "CÔNG TY TNHH NGUYÊN HƯỞNG"),
    "ORIENTAL_MULTIMODAL": ("VN14", "100024393", "Công Ty TNHH Logistics Vận Tải Đa Phương Thức"),
    # 美金 DEBIT NOTE，使用者 2026-11-10 指定獨立供應商代號 VN14_USD (同一家公司，廠商代號相同)
    "ORIENTAL_USD": ("VN14_USD", "100024393", "Công Ty TNHH Logistics Vận Tải Đa Phương Thức"),
    "SONG_TRIEU": ("VN16", "100010180", "CÔNG TY TNHH SONG TRIỀU"),
    "TIMES_CARGO": ("VN17", "100008880", "CÔNG TY TNHH DỊCH VỤ GIAO NHẬN HÀNG HÓA"),
    "SAIGON_TRANSPORT": ("VN18", "100007723", "CÔNG TY TNHH DỊCH VỤ VẬN CHUYỂN SÀI GÒN"),
    "IMX": ("VN19", "100025119", "CÔNG TY TNHH DỊCH VỤ CHUYỂN PHÁT NHANH IMX"),
    # 越南廠商代號：越南進度管控\By Vietnam supplier (Top 20) 的資料夾名稱
    "DSV_VN": ("VN08", "100007866", "CÔNG TY TNHH DSV AIR & SEA"),
    "MOL": ("VN09", "100009335", "CÔNG TY TNHH MOL LOGISTICS VIỆT NAM"),
    "DHL_VN": ("VN13", "100007618", "CTY CP GIAO NHẬN TOÀN CẦU DHL (VIỆTNAM)"),
    "EXPEDITORS": ("VN15", "100008945", "CÔNG TY TNHH EXPEDITORS VIỆT NAM"),
    "YUSEN": ("VN20", "100007349", "CHI NHÁNH CÔNG TY TNHH YUSEN LOGISTICS (VIỆTNAM) TẠI THÀNH PHỐ HỒ CHÍ MINH"),
}

for _key, (_number, _vendor_code, _name) in SUPPLIER_NUMBERS.items():
    set_supplier_number(_key, _number, _vendor_code, _name)

# 越南 TOP20 的英文公司名稱 (使用者附件 "Vietnam company English name.xlsx"，2026-11-04)。
# 輸出 Excel 的「供應商」欄改用英文名稱 (見 extract_utils_streamlit.supplier_label())。
VN_ENGLISH_NAMES = {
    "TRANS_WAGON": "TRANS WAGON INTERNATIONAL CO., LTD",
    "VINH_HUNG_LONG": "VINH HUNG LONG CO., LTD",
    "RONG_A_CHAU": "Asia Dragon Forwarding Trading Company Limited",
    "DHL_EXPRESS_VN": "DHL - VNPT EXPRESS LTD.",
    "CONG_THANH": "CONG THANH COMPANY LTD.",
    "BINH_TAN": "HOP TAC XA VAN TAI DICH VU DU LICH BINH TAN",
    "MAERSK_VN": "MAERSK LOGISTICS & SERVICES VIETNAM COMPANY LIMITED",
    "DSV_VN": "DSV AIR & SEA COMPANY LIMITED",
    "MOL": "MOL LOGISTICS TRANSPORTATION VIETNAM INC",
    "KN_VN": "Kuehne + Nagel Co., Ltd",
    "APL_VN": "APL Logistics Vietnam Co., Ltd",
    "NGUYEN_HUONG": "NGUYEN HUONG COMPANY LTD.",
    "DHL_VN": "DHL GLOBAL FORWARDING (VIETNAM) CORPORATION",
    "ORIENTAL_MULTIMODAL": "ORIENTAL LOGISTICS MULTIMODAL TRANSPORT COMPANY LIMITED.",
    "ORIENTAL_USD": "ORIENTAL LOGISTICS MULTIMODAL TRANSPORT COMPANY LIMITED.",
    "EXPEDITORS": "EXPEDITORS VIETNAM COMPANY LIMITED",
    "SONG_TRIEU": "SONG TRIEU COMPANY LIMITED",
    "TIMES_CARGO": "Times Cargo Logistic Co., Ltd - TCL",
    "SAIGON_TRANSPORT": "SAI GON AIR & OCEAN FREIGHT SCE.CO.LTD",
    "IMX": "IMX EXPRESS SERVICE COMPANY LIMITED",
    "YUSEN": "YUSEN LOGISTICS (VIETNAM) CO., LTD - HO CHI MINH BRANCH",
}
for _key, _english in VN_ENGLISH_NAMES.items():
    SUPPLIER_REGISTRY[_key]["english_name"] = _english

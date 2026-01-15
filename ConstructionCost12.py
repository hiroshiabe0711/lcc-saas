import tkinter as tk
from tkinter import ttk, messagebox
import math
import json
import os
from datetime import datetime, date

# フォント設定
title_font = ("Arial", 14, "bold")
sub_title_font = ("Arial", 12, "bold")
item_font = ("Arial", 11, "bold")
label_font = ("Arial", 10)
sheet_font = ("Arial", 9)
table_font = ("Arial", 8)
tab_font = (None, 8)
# 列幅設定
no_width=5
component_name_width=30
short_name_width=20
man_day_rate_width=15
labor_cost_width=15
sub_materials_width=15
unit_cost_width=15
per_unit_area_width=15
price_rate_width=10
others_width=10



class BuildingCostCalculator:

    def __init__(self, root):

        self.root = root
        self.base_title = "建物コスト算出システム"
        self.root.title(self.base_title)
        self.root.geometry("800x900")


        #アイコンを設定
        try:
            self.root.iconbitmap("icon.ico")  # .icoファイルのパスを指定
        except:
            pass  # アイコンファイルがない場合はデフォルトのまま

        # 物件管理用の変数
        self.current_property_id = None
        self.current_property_name = tk.StringVar(value="新規物件")
        self.data_modified = False
        self.properties_data = {}

        # データ保存先ディレクトリ
        self.data_dir = "onsen_properties"
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)


        # メインコンテナ作成
        main_container = ttk.Frame(root)
        main_container.pack(fill=tk.BOTH, expand=True)


        self.floor_areas = {}   # 動的な値を初期化する為に必要

        # 建設地域掛け率のグローバル変数を初期化
        self.region_rates = {
            '東京都内': tk.StringVar(value="100"),
            'その他関東': tk.StringVar(value="90"),
            '関西': tk.StringVar(value="85"),
            '地方都市': tk.StringVar(value="80"),
            '地方': tk.StringVar(value="75")
        }

        # 建設地域掛け率データファイルのパス
        self.region_rate_file = os.path.join(
            os.path.dirname(__file__),
            'region_rate_data.json'
        )

        # 建設地域掛け率データの読み込み
        loaded_region_data = self.load_region_rate_data()
        if loaded_region_data:
            for region, rate in loaded_region_data.items():
                if region in self.region_rates:
                    self.region_rates[region].set(str(rate))

        # 設定用の変数(初期値)
        self.lcc_building_area = tk.StringVar(value="0")
        self.cost_summary_total_area_var = tk.StringVar(value="合計: 0 ㎡")#合計面積表示用変数
        self.selected_spec_grade = tk.StringVar(value="Gensen")
        self.lcc_building_area.trace_add("write", self.update_construction_unit_price)# 建物延床面積が変更されたときに単価計算を呼び出すバインド
        self.new_construction_cost = tk.StringVar(value="0")  # 新築工事推定金額


        self.region = tk.StringVar()#建設地域
        self.base_date = tk.StringVar()  # 基準日
        self.date_difference = tk.IntVar(value=0)  # 日数差
        self.completion_date = tk.StringVar()#建物竣工年 (yyyy/mm/dd)
        self.floor_area = tk.StringVar()#延床面積 (㎡)
        self.structure_type = tk.StringVar()#建物構造
        self.floors = tk.StringVar()#地上階数
        self.underground_floors = tk.StringVar()#地下階数
        self.spec_grade = tk.StringVar()#仕様・グレード

        # 設定用の変数(初期値)
        self.misc_cost_rate = tk.DoubleVar(value=0.15)  # 諸経費率 15%

        # 仕様別の割増率設定
        self.spec1_rate = tk.DoubleVar(value=0.10)  # 仕様1: 10%
        self.spec2_rate = tk.DoubleVar(value=0.25)  # 仕様2: 25%
        self.spec3_rate = tk.DoubleVar(value=0.45)  # 仕様3: 45%

        # 付加設備のチェックボックス用変数
        self.heat_source = tk.BooleanVar()
        self.elevator = tk.BooleanVar()
        self.gas_equipment = tk.BooleanVar()
        self.septic_tank = tk.BooleanVar()
        self.central_ac = tk.BooleanVar()  # 油送設備
        self.fire_safety = tk.BooleanVar()  # 受変電設備
        self.generator = tk.BooleanVar()  # 自家発電設備

        # 付加設備の金額表示用変数
        self.heat_source_cost = tk.StringVar(value="0")
        self.heat_source_process = tk.StringVar(value="")
        self.heat_source_fuel_type = tk.StringVar(value="ガス")  # ← 追加：燃料種別
        self.heat_source_unit_count = tk.StringVar(value="1")  # ← 追加：台数
        self.heat_source_required_output = tk.StringVar(value="0")  # ← 追加：必要出力表示用
        self.elevator_cost = tk.StringVar(value="0")
        self.gas_equipment_cost = tk.StringVar(value="0")
        self.septic_tank_cost = tk.StringVar(value="0")
        self.oil_supply_cost = tk.StringVar(value="0")
        self.substation_cost = tk.StringVar(value="0")
        self.generator_cost = tk.StringVar(value="0")

        # 付加設備の計算過程表示用変数
        self.elevator_process = tk.StringVar(value="")
        self.gas_equipment_process = tk.StringVar(value="")
        self.septic_tank_process = tk.StringVar(value="")
        self.oil_supply_process = tk.StringVar(value="")
        self.substation_process = tk.StringVar(value="")
        self.generator_process = tk.StringVar(value="")

        # 付加設備小計
        self.additional_equipment_subtotal = tk.StringVar(value="0")


        # LCC分析タブ用の変数
        self.lcc_building_area = tk.StringVar(value="0")  # 建物面積
        self.lcc_construction_unit_price = tk.StringVar(value="0")  # 建設基準単価  円/m2
        self.lcc_region_rate_display = tk.StringVar(value="1.00")
        self.direct_construction_cost_rate = tk.StringVar(value="0.00")  # 直接工事費率
        self.direct_construction_cost_amount = tk.StringVar(value="0")  # 直接工事費金額
        self.net_construction_cost_rate = tk.StringVar(value="0")  # 純工事費（自動計算）
        self.net_construction_cost_amount = tk.StringVar(value="0")  # 純工事費金額
        self.common_temporary_cost_rate = tk.StringVar(value="4")  # 共通仮設費 4%
        self.common_temporary_cost_amount = tk.StringVar(value="0")  # 共通仮設費金額
        self.general_admin_fee_rate = tk.StringVar(value="12")  # 一般管理費 12%
        self.general_admin_fee_amount = tk.StringVar(value="0")  # 一般管理費等金額
        self.site_management_cost_rate = tk.StringVar(value="7")  # 現場管理費 7%
        self.site_management_cost_amount = tk.StringVar(value="0")  # 現場管理費金額


        # 工事価格比率
        self.lcc_construction_material_rate = tk.StringVar(value="40.0")  # 建築工事材(建築)

        # 各種設備工事の比率(直接工事費に対する%)
        self.lcc_equipment_general_rate = tk.StringVar(value="15")  # 空調機器設備
        self.lcc_equipment_thermal1_rate = tk.StringVar(value="10")  # 熱源配管設備
        self.lcc_equipment_thermal2_rate = tk.StringVar(value="25")  # 空調機器設備
        self.lcc_equipment_thermal3_rate = tk.StringVar(value="13")  # 空調配管設備
        self.lcc_equipment_thermal4_rate = tk.StringVar(value="8")  # 換気機器設備
        self.lcc_equipment_thermal5_rate = tk.StringVar(value="6")  # 換気ダクト設備
        self.lcc_equipment_thermal6_rate = tk.StringVar(value="15")  # 自動制御設備
        self.lcc_equipment_thermal7_rate = tk.StringVar(value="8")  # 排煙設備

        # 衛生設備(工事価格の12%)
        self.lcc_sanitary_general_rate = tk.StringVar(value="6")  # 衛生器具設備
        self.lcc_sanitary_water_rate = tk.StringVar(value="6")  # 給水機器設備
        self.lcc_sanitary_drainage_rate = tk.StringVar(value="7")  # 給水配管設備
        self.lcc_sanitary_hotwater_rate = tk.StringVar(value="9")  # 給湯機器設備
        self.lcc_sanitary_gas_rate = tk.StringVar(value="7")  # 給湯配管設備
        self.lcc_sanitary_drainage2_rate = tk.StringVar(value="16")  # 排水設備
        self.lcc_sanitary_gas2_rate = tk.StringVar(value="6")  # ガス設備
        self.lcc_sanitary_kitchen_rate = tk.StringVar(value="5")  # 厨房設備
        self.lcc_sanitary_purification_rate = tk.StringVar(value="8")  # 浄化槽設備
        self.lcc_sanitary_fire_rate = tk.StringVar(value="28")  # 消火設備
        self.lcc_sanitary_drainage3_rate = tk.StringVar(value="2")  # 雑排水設備

        # 電気設備(工事価格の15.5%)
        self.lcc_electric_reception_rate = tk.StringVar(value="10")  # 受変電設備
        self.lcc_electric_power_rate = tk.StringVar(value="8")  # 自家発電設備
        self.lcc_electric_dynamic_rate = tk.StringVar(value="20")  # 動力設備
        self.lcc_electric_lighting_rate = tk.StringVar(value="20")  # 電灯設備
        self.lcc_electric_info_rate = tk.StringVar(value="15")  # 情報通信設備
        self.lcc_electric_monitor_rate = tk.StringVar(value="12")  # 中央監視制御設備
        self.lcc_electric_fire_rate = tk.StringVar(value="15")  # 火災警報設備

        # 昇降機設備
        self.lcc_elevator_unit_price = tk.StringVar(value="42500000")  # @42,500,000 1基/100室
        self.lcc_elevator_count = tk.StringVar(value="1")  # エレベーター台数

        # ダムウェーター
        self.lcc_dumbwaiter_unit_price = tk.StringVar(value="3000000")  # @3,000,000 老数(厨房接続)1
        self.lcc_dumbwaiter_count = tk.StringVar(value="0")  # ダムウェーター台数

        # LCC期間設定
        self.lcc_analysis_years = tk.StringVar(value="70")  # 分析期間
        self.lcc_discount_rate = tk.StringVar(value="3.0")  # 割引率

        # 維持管理費(建設費の年間%)
        self.lcc_annual_maintenance_rate = tk.StringVar(value="1.5")  # 年間維持管理費率 1.5%

        # 大規模修繕(建設費の%)
        self.lcc_major_repair_rate = tk.StringVar(value="20")  # 大規模修繕費率 20%
        self.lcc_major_repair_interval = tk.StringVar(value="12")  # 修繕周期 12年

        # 運用費
        self.lcc_energy_unit_cost = tk.StringVar(value="550")  # エネルギー単価 円/m2/月
        self.lcc_water_unit_cost = tk.StringVar(value="55")  # 水道単価 円/m2/月
        self.lcc_insurance_rate = tk.StringVar(value="0.15")  # 保険料率 0.15%
        self.lcc_tax_rate = tk.StringVar(value="1.4")  # 固定資産税率 1.4%

        # 運営費関連の変数
        self.lcc_personnel_cost = tk.StringVar(value="0")  # 人件費
        self.lcc_amenity_cost = tk.StringVar(value="0")  # アメニティ費
        self.lcc_ota_ad_cost = tk.StringVar(value="0")  # OTA・広告費
        self.lcc_food_cost = tk.StringVar(value="0")  # 食材費
        self.lcc_linen_cost = tk.StringVar(value="0")  # リネン費
        self.lcc_operation_cost = tk.StringVar(value="0")  # 運営費（合計、読み取り専用）


        # 温泉設備の設定変数
        self.bath_open_hours = tk.DoubleVar(value=6.0)#浴室営業時間
        self.wash_time = tk.DoubleVar(value=10.0)#洗い場利用時間
        self.bath_utilization_rate = tk.DoubleVar(value=100.0)#洗場利用率
        self.tub_time = tk.DoubleVar(value=20.0)#浴槽利用時間
        self.changing_room_time = tk.DoubleVar(value=20.0)#脱衣室利用時間
        self.men_face_wash_time = tk.DoubleVar(value=10.0)#男洗面器利用時間
        self.women_face_wash_time = tk.DoubleVar(value=15.0)#女洗面器利用時間
        self.peak_factor = tk.DoubleVar(value=4.0)#ピーク係数
        self.men_ratio = tk.DoubleVar(value=55.0)#利用者男性率
        self.women_ratio = tk.DoubleVar(value=45.0)#利用者女性率
        self.wash_area_per_person = tk.DoubleVar(value=1.1)#洗い場使用面積
        self.changing_room_per_person = tk.DoubleVar(value=1.1)#脱衣室使用面積
        self.men_changing_room_add = tk.DoubleVar(value=150.0)#男脱衣室面積の割増
        self.women_changing_room_add = tk.DoubleVar(value=175.0)#女脱衣室面積の割増
        self.bath_hallway_add = tk.DoubleVar(value=200.0)#浴室内部通路割増
        self.tub_area_per_person = tk.DoubleVar(value=1.2)#浴槽使用面積
        self.bath_area_per_person = tk.DoubleVar(value=4.0)#浴室使用面積
        self.bath_area_add = tk.DoubleVar(value=120.0)#浴室面積の割増
        self.tub_depth = tk.DoubleVar(value=0.7)#浴槽深さ
        self.face_wash_utilization_rate = tk.DoubleVar(value=40.0)#洗面器利用率
        self.faucet_utilization_rate = tk.DoubleVar(value=50.0)#カラン利用率
        self.onsen_subtotal = tk.StringVar(value='0')
        # 温泉設備タブの建築項目用変数(男性用・女性用)を追加
        self.men_indoor_bath_check = tk.BooleanVar(value=False)
        self.men_indoor_bath_area = tk.StringVar(value="0")
        self.men_indoor_bath_cost = tk.StringVar(value="0")
        self.men_outdoor_bath_check = tk.BooleanVar(value=False)
        self.men_outdoor_bath_area = tk.StringVar(value="0")
        self.men_outdoor_bath_cost = tk.StringVar(value="0")
        self.men_bathroom_check = tk.BooleanVar(value=False)
        self.men_bathroom_area = tk.StringVar(value="0")
        self.men_bathroom_cost = tk.StringVar(value="0")
        self.men_bath_arch_process = tk.StringVar(value="")

        self.women_indoor_bath_check = tk.BooleanVar(value=False)
        self.women_indoor_bath_area = tk.StringVar(value="0")
        self.women_indoor_bath_cost = tk.StringVar(value="0")
        self.women_outdoor_bath_check = tk.BooleanVar(value=False)
        self.women_outdoor_bath_area = tk.StringVar(value="0")
        self.women_outdoor_bath_cost = tk.StringVar(value="0")
        self.women_bathroom_check = tk.BooleanVar(value=False)
        self.women_bathroom_area = tk.StringVar(value="0")
        self.women_bathroom_cost = tk.StringVar(value="0")
        self.women_bath_arch_process = tk.StringVar(value="")
        # 男性用
        self.men_partition_count = tk.StringVar(value="0")
        self.men_partition_cost = tk.StringVar(value="0")
        self.men_partition_recommended = tk.StringVar(value="")
        self.men_wash_lighting_count = tk.StringVar(value="0")
        self.men_wash_lighting_cost = tk.StringVar(value="0")
        self.men_wash_lighting_recommended = tk.StringVar(value="")
        self.men_shower_faucet_count = tk.StringVar(value="0")
        self.men_shower_faucet_cost = tk.StringVar(value="0")
        self.men_shower_faucet_recommended = tk.StringVar(value="")
        # 女性用
        self.women_partition_count = tk.StringVar(value="0")
        self.women_partition_cost = tk.StringVar(value="0")
        self.women_partition_recommended = tk.StringVar(value="")
        self.women_wash_lighting_count = tk.StringVar(value="0")
        self.women_wash_lighting_cost = tk.StringVar(value="0")
        self.women_wash_lighting_recommended = tk.StringVar(value="")
        self.women_shower_faucet_count = tk.StringVar(value="0")
        self.women_shower_faucet_cost = tk.StringVar(value="0")
        self.women_shower_faucet_recommended = tk.StringVar(value="")
        # 洗い場設備の計算過程表示用
        self.men_wash_equipment_process = tk.StringVar(value="")
        self.women_wash_equipment_process = tk.StringVar(value="")

        # チェックボックス用変数
        self.men_partition_check = tk.BooleanVar(value=False)
        self.men_wash_lighting_check = tk.BooleanVar(value=False)
        self.men_shower_faucet_check = tk.BooleanVar(value=False)
        self.women_partition_check = tk.BooleanVar(value=False)
        self.women_wash_lighting_check = tk.BooleanVar(value=False)
        self.women_shower_faucet_check = tk.BooleanVar(value=False)
        # 特殊設備のチェックボックス
        self.men_sauna = tk.BooleanVar(value=False)
        self.men_jacuzzi = tk.BooleanVar(value=False)
        self.men_water_bath = tk.BooleanVar(value=False)
        self.women_sauna = tk.BooleanVar(value=False)
        self.women_jacuzzi = tk.BooleanVar(value=False)
        self.women_water_bath = tk.BooleanVar(value=False)
        # 特殊設備の金額変数
        self.men_sauna_cost = tk.StringVar(value="0")
        self.men_jacuzzi_cost = tk.StringVar(value="0")
        self.men_water_bath_cost = tk.StringVar(value="0")
        self.women_sauna_cost = tk.StringVar(value="0")
        self.women_jacuzzi_cost = tk.StringVar(value="0")
        self.women_water_bath_cost = tk.StringVar(value="0")


        # レストランタブの設定変数
        # 建築
        self.restaurant_arch_restaurant_check = tk.BooleanVar(value=False)
        self.restaurant_arch_restaurant_area = tk.StringVar(value="0")
        self.restaurant_arch_restaurant_cost = tk.StringVar(value="0")
        self.restaurant_arch_livekitchen_check = tk.BooleanVar(value=False)
        self.restaurant_arch_livekitchen_area = tk.StringVar(value="0")
        self.restaurant_arch_livekitchen_cost = tk.StringVar(value="0")
        self.restaurant_arch_kitchen_check = tk.BooleanVar(value=False)
        self.restaurant_arch_kitchen_area = tk.StringVar(value="0")
        self.restaurant_arch_kitchen_cost = tk.StringVar(value="0")

        # 機械設備
        self.restaurant_mech_ac_check = tk.BooleanVar(value=False)
        self.restaurant_mech_ac_count = tk.StringVar(value="0")
        self.restaurant_mech_ac_cost = tk.StringVar(value="0")
        self.restaurant_mech_sprinkler_check = tk.BooleanVar(value=False)
        self.restaurant_mech_sprinkler_count = tk.StringVar(value="0")
        self.restaurant_mech_sprinkler_cost = tk.StringVar(value="0")
        self.restaurant_mech_fire_hose_check = tk.BooleanVar(value=False)
        self.restaurant_mech_fire_hose_count = tk.StringVar(value="0")
        self.restaurant_mech_fire_hose_cost = tk.StringVar(value="0")
        self.restaurant_mech_hood_check = tk.BooleanVar(value=False)
        self.restaurant_mech_hood_count = tk.StringVar(value="0")
        self.restaurant_mech_hood_cost = tk.StringVar(value="0")

        # 電気設備
        self.restaurant_elec_led_check = tk.BooleanVar(value=False)
        self.restaurant_elec_led_count = tk.StringVar(value="0")
        self.restaurant_elec_led_cost = tk.StringVar(value="0")
        self.restaurant_elec_smoke_check = tk.BooleanVar(value=False)
        self.restaurant_elec_smoke_count = tk.StringVar(value="0")
        self.restaurant_elec_smoke_cost = tk.StringVar(value="0")
        self.restaurant_elec_exit_light_check = tk.BooleanVar(value=False)
        self.restaurant_elec_exit_light_count = tk.StringVar(value="0")
        self.restaurant_elec_exit_light_cost = tk.StringVar(value="0")
        self.restaurant_elec_speaker_check = tk.BooleanVar(value=False)
        self.restaurant_elec_speaker_count = tk.StringVar(value="0")
        self.restaurant_elec_speaker_cost = tk.StringVar(value="0")
        self.restaurant_elec_outlet_check = tk.BooleanVar(value=False)
        self.restaurant_elec_outlet_count = tk.StringVar(value="0")
        self.restaurant_elec_outlet_cost = tk.StringVar(value="0")

        # 厨房設備（動的に厨房機器設定から生成）
        self.restaurant_kitchen_equipment_checks = []
        self.restaurant_kitchen_equipment_counts = []
        self.restaurant_kitchen_equipment_costs = []

        # 家具
        self.restaurant_furniture_bar_counter_check = tk.BooleanVar(value=False)
        self.restaurant_furniture_bar_counter_count = tk.StringVar(value="0")
        self.restaurant_furniture_bar_counter_cost = tk.StringVar(value="0")
        self.restaurant_furniture_soft_drink_counter_check = tk.BooleanVar(value=False)
        self.restaurant_furniture_soft_drink_counter_count = tk.StringVar(value="0")
        self.restaurant_furniture_soft_drink_counter_cost = tk.StringVar(value="0")
        self.restaurant_furniture_alcohol_counter_check = tk.BooleanVar(value=False)
        self.restaurant_furniture_alcohol_counter_count = tk.StringVar(value="0")
        self.restaurant_furniture_alcohol_counter_cost = tk.StringVar(value="0")
        self.restaurant_furniture_cutlery_counter_check = tk.BooleanVar(value=False)
        self.restaurant_furniture_cutlery_counter_count = tk.StringVar(value="0")
        self.restaurant_furniture_cutlery_counter_cost = tk.StringVar(value="0")
        self.restaurant_furniture_ice_counter_check = tk.BooleanVar(value=False)
        self.restaurant_furniture_ice_counter_count = tk.StringVar(value="0")
        self.restaurant_furniture_ice_counter_cost = tk.StringVar(value="0")
        self.restaurant_furniture_soft_cream_counter_check = tk.BooleanVar(value=False)
        self.restaurant_furniture_soft_cream_counter_count = tk.StringVar(value="0")
        self.restaurant_furniture_soft_cream_counter_cost = tk.StringVar(value="0")
        self.restaurant_furniture_return_counter_check = tk.BooleanVar(value=False)
        self.restaurant_furniture_return_counter_count = tk.StringVar(value="0")
        self.restaurant_furniture_return_counter_cost = tk.StringVar(value="0")

        # 推奨値表示用
        self.restaurant_mech_ac_recommended = tk.StringVar(value="")
        self.restaurant_mech_sprinkler_recommended = tk.StringVar(value="")
        self.restaurant_mech_fire_hose_recommended = tk.StringVar(value="")
        self.restaurant_mech_hood_recommended = tk.StringVar(value="")
        self.restaurant_elec_led_recommended = tk.StringVar(value="")
        self.restaurant_elec_smoke_recommended = tk.StringVar(value="")
        self.restaurant_elec_exit_light_recommended = tk.StringVar(value="")
        self.restaurant_elec_speaker_recommended = tk.StringVar(value="")
        self.restaurant_elec_outlet_recommended = tk.StringVar(value="")
        self.restaurant_furniture_bar_counter_recommended = tk.StringVar(value="")
        self.restaurant_furniture_soft_drink_counter_recommended = tk.StringVar(value="")
        self.restaurant_furniture_alcohol_counter_recommended = tk.StringVar(value="")
        self.restaurant_furniture_cutlery_counter_recommended = tk.StringVar(value="")
        self.restaurant_furniture_ice_counter_recommended = tk.StringVar(value="")
        self.restaurant_furniture_soft_cream_counter_recommended = tk.StringVar(value="")
        self.restaurant_furniture_return_counter_recommended = tk.StringVar(value="")

        # 計算過程表示用
        self.restaurant_arch_restaurant_process = tk.StringVar(value="")
        self.restaurant_arch_livekitchen_process = tk.StringVar(value="")
        self.restaurant_arch_kitchen_process = tk.StringVar(value="")
        self.restaurant_mech_process = tk.StringVar(value="")
        self.restaurant_elec_process = tk.StringVar(value="")
        self.restaurant_kitchen_process = tk.StringVar(value="")
        self.restaurant_furniture_process = tk.StringVar(value="")

        # ラウンジタブの設定変数
        self.lounge_arch_front_check = tk.BooleanVar(value=False)
        self.lounge_arch_front_area = tk.StringVar(value="0")
        self.lounge_arch_front_cost = tk.StringVar(value="0")
        self.lounge_arch_lobby_check = tk.BooleanVar(value=False)
        self.lounge_arch_lobby_area = tk.StringVar(value="0")
        self.lounge_arch_lobby_cost = tk.StringVar(value="0")
        self.lounge_arch_facade_check = tk.BooleanVar(value=False)
        self.lounge_arch_facade_area = tk.StringVar(value="0")
        self.lounge_arch_facade_cost = tk.StringVar(value="0")
        self.lounge_arch_shop_check = tk.BooleanVar(value=False)
        self.lounge_arch_shop_area = tk.StringVar(value="0")
        self.lounge_arch_shop_cost = tk.StringVar(value="0")

        self.lounge_mech_ac_check = tk.BooleanVar(value=False)
        self.lounge_mech_ac_count = tk.StringVar(value="0")
        self.lounge_mech_ac_cost = tk.StringVar(value="0")
        self.lounge_mech_sprinkler_check = tk.BooleanVar(value=False)
        self.lounge_mech_sprinkler_count = tk.StringVar(value="0")
        self.lounge_mech_sprinkler_cost = tk.StringVar(value="0")
        self.lounge_mech_fire_hose_check = tk.BooleanVar(value=False)
        self.lounge_mech_fire_hose_count = tk.StringVar(value="0")
        self.lounge_mech_fire_hose_cost = tk.StringVar(value="0")

        self.lounge_elec_led_check = tk.BooleanVar(value=False)
        self.lounge_elec_led_count = tk.StringVar(value="0")
        self.lounge_elec_led_cost = tk.StringVar(value="0")
        self.lounge_elec_smoke_check = tk.BooleanVar(value=False)
        self.lounge_elec_smoke_count = tk.StringVar(value="0")
        self.lounge_elec_smoke_cost = tk.StringVar(value="0")
        self.lounge_elec_exit_light_check = tk.BooleanVar(value=False)
        self.lounge_elec_exit_light_count = tk.StringVar(value="0")
        self.lounge_elec_exit_light_cost = tk.StringVar(value="0")
        self.lounge_elec_speaker_check = tk.BooleanVar(value=False)
        self.lounge_elec_speaker_count = tk.StringVar(value="0")
        self.lounge_elec_speaker_cost = tk.StringVar(value="0")
        self.lounge_elec_outlet_check = tk.BooleanVar(value=False)
        self.lounge_elec_outlet_count = tk.StringVar(value="0")
        self.lounge_elec_outlet_cost = tk.StringVar(value="0")

        self.lounge_arch_front_process = tk.StringVar(value="")
        self.lounge_arch_lobby_process = tk.StringVar(value="")
        self.lounge_arch_facade_process = tk.StringVar(value="")
        self.lounge_arch_shop_process = tk.StringVar(value="")
        self.lounge_mech_process = tk.StringVar(value="")
        self.lounge_elec_process = tk.StringVar(value="")

        self.lounge_mech_ac_recommended = tk.StringVar(value="")
        self.lounge_mech_sprinkler_recommended = tk.StringVar(value="")
        self.lounge_mech_fire_hose_recommended = tk.StringVar(value="")
        self.lounge_elec_led_recommended = tk.StringVar(value="")
        self.lounge_elec_smoke_recommended = tk.StringVar(value="")
        self.lounge_elec_exit_light_recommended = tk.StringVar(value="")
        self.lounge_elec_speaker_recommended = tk.StringVar(value="")
        self.lounge_elec_outlet_recommended = tk.StringVar(value="")

        # アミューズメントタブの設定変数
        self.amusement_arch_pingpong_check = tk.BooleanVar(value=False)
        self.amusement_arch_pingpong_area = tk.StringVar(value="0")
        self.amusement_arch_pingpong_cost = tk.StringVar(value="0")
        self.amusement_arch_kids_check = tk.BooleanVar(value=False)
        self.amusement_arch_kids_area = tk.StringVar(value="0")
        self.amusement_arch_kids_cost = tk.StringVar(value="0")
        self.amusement_arch_manga_check = tk.BooleanVar(value=False)
        self.amusement_arch_manga_area = tk.StringVar(value="0")
        self.amusement_arch_manga_cost = tk.StringVar(value="0")

        self.amusement_mech_ac_check = tk.BooleanVar(value=False)
        self.amusement_mech_ac_count = tk.StringVar(value="0")
        self.amusement_mech_ac_cost = tk.StringVar(value="0")
        self.amusement_mech_sprinkler_check = tk.BooleanVar(value=False)
        self.amusement_mech_sprinkler_count = tk.StringVar(value="0")
        self.amusement_mech_sprinkler_cost = tk.StringVar(value="0")
        self.amusement_mech_fire_hose_check = tk.BooleanVar(value=False)
        self.amusement_mech_fire_hose_count = tk.StringVar(value="0")
        self.amusement_mech_fire_hose_cost = tk.StringVar(value="0")

        self.amusement_elec_led_check = tk.BooleanVar(value=False)
        self.amusement_elec_led_count = tk.StringVar(value="0")
        self.amusement_elec_led_cost = tk.StringVar(value="0")
        self.amusement_elec_smoke_check = tk.BooleanVar(value=False)
        self.amusement_elec_smoke_count = tk.StringVar(value="0")
        self.amusement_elec_smoke_cost = tk.StringVar(value="0")
        self.amusement_elec_exit_light_check = tk.BooleanVar(value=False)
        self.amusement_elec_exit_light_count = tk.StringVar(value="0")
        self.amusement_elec_exit_light_cost = tk.StringVar(value="0")
        self.amusement_elec_speaker_check = tk.BooleanVar(value=False)
        self.amusement_elec_speaker_count = tk.StringVar(value="0")
        self.amusement_elec_speaker_cost = tk.StringVar(value="0")
        self.amusement_elec_outlet_check = tk.BooleanVar(value=False)
        self.amusement_elec_outlet_count = tk.StringVar(value="0")
        self.amusement_elec_outlet_cost = tk.StringVar(value="0")

        self.amusement_arch_pingpong_process = tk.StringVar(value="")
        self.amusement_arch_kids_process = tk.StringVar(value="")
        self.amusement_arch_manga_process = tk.StringVar(value="")
        self.amusement_mech_process = tk.StringVar(value="")
        self.amusement_elec_process = tk.StringVar(value="")

        self.amusement_mech_ac_recommended = tk.StringVar(value="")
        self.amusement_mech_sprinkler_recommended = tk.StringVar(value="")
        self.amusement_mech_fire_hose_recommended = tk.StringVar(value="")
        self.amusement_elec_led_recommended = tk.StringVar(value="")
        self.amusement_elec_smoke_recommended = tk.StringVar(value="")
        self.amusement_elec_exit_light_recommended = tk.StringVar(value="")
        self.amusement_elec_speaker_recommended = tk.StringVar(value="")
        self.amusement_elec_outlet_recommended = tk.StringVar(value="")

        # 通路の機械設備変数
        self.hallway_mech_ac_check = tk.BooleanVar(value=False)
        self.hallway_mech_ac_count = tk.StringVar(value="0")
        self.hallway_mech_ac_cost = tk.StringVar(value="0")
        self.hallway_mech_sprinkler_check = tk.BooleanVar(value=False)
        self.hallway_mech_sprinkler_count = tk.StringVar(value="0")
        self.hallway_mech_sprinkler_cost = tk.StringVar(value="0")
        self.hallway_mech_fire_hose_check = tk.BooleanVar(value=False)
        self.hallway_mech_fire_hose_count = tk.StringVar(value="0")
        self.hallway_mech_fire_hose_cost = tk.StringVar(value="0")

        # 通路の電気設備変数
        self.hallway_elec_led_check = tk.BooleanVar(value=False)
        self.hallway_elec_led_count = tk.StringVar(value="0")
        self.hallway_elec_led_cost = tk.StringVar(value="0")
        self.hallway_elec_smoke_check = tk.BooleanVar(value=False)
        self.hallway_elec_smoke_count = tk.StringVar(value="0")
        self.hallway_elec_smoke_cost = tk.StringVar(value="0")
        self.hallway_elec_exit_light_check = tk.BooleanVar(value=False)
        self.hallway_elec_exit_light_count = tk.StringVar(value="0")
        self.hallway_elec_exit_light_cost = tk.StringVar(value="0")
        self.hallway_elec_speaker_check = tk.BooleanVar(value=False)
        self.hallway_elec_speaker_count = tk.StringVar(value="0")
        self.hallway_elec_speaker_cost = tk.StringVar(value="0")
        self.hallway_elec_outlet_check = tk.BooleanVar(value=False)
        self.hallway_elec_outlet_count = tk.StringVar(value="0")
        self.hallway_elec_outlet_cost = tk.StringVar(value="0")

        # 通路の計算過程表示用
        self.hallway_mech_process = tk.StringVar(value="")
        self.hallway_elec_process = tk.StringVar(value="")

        # 通路の推奨個数表示用
        self.hallway_mech_ac_recommended = tk.StringVar(value="")
        self.hallway_mech_sprinkler_recommended = tk.StringVar(value="")
        self.hallway_mech_fire_hose_recommended = tk.StringVar(value="")
        self.hallway_elec_led_recommended = tk.StringVar(value="")
        self.hallway_elec_smoke_recommended = tk.StringVar(value="")
        self.hallway_elec_exit_light_recommended = tk.StringVar(value="")
        self.hallway_elec_speaker_recommended = tk.StringVar(value="")
        self.hallway_elec_outlet_recommended = tk.StringVar(value="")

        # 客室タブの設定変数
        self.capacity_japanese_room = tk.StringVar(value="4.5")
        self.capacity_japanese_western_room = tk.StringVar(value="4.0")
        self.capacity_japanese_bed_room = tk.StringVar(value="3.0")
        self.capacity_western_room = tk.StringVar(value="3.0")

        # 客室建築項目
        self.guest_arch_japanese_check = tk.BooleanVar(value=False)
        self.guest_arch_japanese_area = tk.StringVar(value="0")
        self.guest_arch_japanese_cost = tk.StringVar(value="0")
        self.guest_arch_japanese_western_check = tk.BooleanVar(value=False)
        self.guest_arch_japanese_western_area = tk.StringVar(value="0")
        self.guest_arch_japanese_western_cost = tk.StringVar(value="0")
        self.guest_arch_japanese_bed_check = tk.BooleanVar(value=False)
        self.guest_arch_japanese_bed_area = tk.StringVar(value="0")
        self.guest_arch_japanese_bed_cost = tk.StringVar(value="0")
        self.guest_arch_western_check = tk.BooleanVar(value=False)
        self.guest_arch_western_area = tk.StringVar(value="0")
        self.guest_arch_western_cost = tk.StringVar(value="0")

        # 客室機械設備
        self.guest_mech_ac_check = tk.BooleanVar(value=False)  # 個別エアコン
        self.guest_mech_ac_count = tk.StringVar(value="0")
        self.guest_mech_ac_cost = tk.StringVar(value="0")
        self.guest_mech_wash_basin_check = tk.BooleanVar(value=False)  # 洗面器（旧sprinkler）
        self.guest_mech_wash_basin_count = tk.StringVar(value="0")
        self.guest_mech_wash_basin_cost = tk.StringVar(value="0")
        self.guest_mech_sprinkler_check = tk.BooleanVar(value=False)  # スプリンクラー（旧fire_hose）
        self.guest_mech_sprinkler_count = tk.StringVar(value="0")
        self.guest_mech_sprinkler_cost = tk.StringVar(value="0")

        # 客室電気設備
        self.guest_elec_main_light_check = tk.BooleanVar(value=False)  # 主室照明（旧led）
        self.guest_elec_main_light_count = tk.StringVar(value="0")
        self.guest_elec_main_light_cost = tk.StringVar(value="0")
        self.guest_elec_smoke_check = tk.BooleanVar(value=False)  # 煙感知器（変更なし）
        self.guest_elec_smoke_count = tk.StringVar(value="0")
        self.guest_elec_smoke_cost = tk.StringVar(value="0")
        self.guest_elec_heat_detector_check = tk.BooleanVar(value=False)  # 熱感知器（旧exit_light）
        self.guest_elec_heat_detector_count = tk.StringVar(value="0")
        self.guest_elec_heat_detector_cost = tk.StringVar(value="0")
        self.guest_elec_speaker_check = tk.BooleanVar(value=False)
        self.guest_elec_speaker_count = tk.StringVar(value="0")
        self.guest_elec_speaker_cost = tk.StringVar(value="0")
        self.guest_elec_outlet_check = tk.BooleanVar(value=False)
        self.guest_elec_outlet_count = tk.StringVar(value="0")
        self.guest_elec_outlet_cost = tk.StringVar(value="0")

        # 客室計算過程表示用
        self.guest_arch_japanese_process = tk.StringVar(value="")
        self.guest_arch_japanese_western_process = tk.StringVar(value="")
        self.guest_arch_japanese_bed_process = tk.StringVar(value="")
        self.guest_arch_western_process = tk.StringVar(value="")

        self.guest_mech_process = tk.StringVar(value="")
        self.guest_elec_process = tk.StringVar(value="")

        # 客室推奨個数表示用
        self.guest_mech_ac_recommended = tk.StringVar(value="")
        self.guest_mech_wash_basin_recommended = tk.StringVar(value="")
        self.guest_mech_sprinkler_recommended = tk.StringVar(value="")
        self.guest_elec_main_light_recommended = tk.StringVar(value="")
        self.guest_elec_smoke_recommended = tk.StringVar(value="")
        self.guest_elec_heat_detector_recommended = tk.StringVar(value="")
        self.guest_elec_speaker_recommended = tk.StringVar(value="")
        self.guest_elec_outlet_recommended = tk.StringVar(value="")

        # 共通客室設備
        self.guest_room_tv_cabinet = tk.BooleanVar(value=False)
        self.guest_room_headboard = tk.BooleanVar(value=False)
        self.guest_private_open_bath = tk.BooleanVar(value=False)

        # 小計用の変数
        self.onsen_subtotal = tk.StringVar(value="0")  # 温泉設備小計用の変数
        self.restaurant_subtotal = tk.StringVar(value="0") # レストラン小計用の変数
        self.lounge_subtotal = tk.StringVar(value="0") # ラウンジ小計用の変数
        self.amusement_subtotal = tk.StringVar(value="0") # アミューズメント小計用の変数
        self.hallway_subtotal = tk.StringVar(value="0")  # 通路小計用の変数
        self.guest_subtotal = tk.StringVar(value="0") # 客室小計用の変数

        # 動的な値を初期化する為に必要
        self.floor_area = {}
        self.hallway_arch_checks = {}
        self.hallway_arch_areas = {}
        self.hallway_arch_costs = {}
        self.hallway_arch_processes = {}

        # 【ここから追加】単価データ参照用のインデックスマップ
        self.structure_index_map = {
            "木造": 0, "RC造": 1, "鉄骨造": 2, "SRC造": 3
        }
        self.grade_index_map = {
            "Gensen": 0, "Premium 1": 1, "Premium 2": 2,
            "TAOYA 1": 3, "TAOYA 2": 4
        }

        # 家具設定単価1のデータ
        self.furniture1_data_file = "furniture1_data.json"
        initial_furniture1_data = [
            (1, "メラミンカウンターA  1800×950", "バーカウンター", 0.1, 27000, 2000, 389000, 0.032),
            (2, "メラミンカウンターB  1800×950", "ソフトドリンクカウンター", 0.1, 27000, 2000, 480000, 0.002),
            (3, "メラミンカウンターC  1800×950", "アルコールカウンター", 0.1, 27000, 2000, 480000, 0.002),
            (4, "メラミンカウンターD  1800×950", "カトラリ―カウンター", 0.1, 27000, 2000, 389000, 0.004),
            (5, "メラミンカウンターE  950×950", "アイスカウンター", 0.1, 27000, 2000, 216000, 0.002),
            (6, "メラミンカウンターE  1350×1350", "ソフトクリームカウンター", 0.1, 27000, 2000, 260000, 0.002),
            (7, "返却台", "返却台", 0.1, 27000, 2000, 30000, 0.02),
            (8, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02),
            (9, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02),
            (10, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02),
            (11, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02),
            (12, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02),
            (13, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02),
            (14, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02),
            (15, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02),
            (16, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02),
            (17, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02),
            (18, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02),
            (19, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02),
            (20, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02),

        ]

        self.initial_furniture1_data = initial_furniture1_data
        self.furniture1_unit_data = []

        loaded_furniture1 = self.load_furniture1_data()
        if loaded_furniture1:
            for data in loaded_furniture1:
                no, name, abbrev, install_hours, labor_cost, misc_material, new_equipment_cost, area_ratio = data
                no_var = tk.StringVar(value=str(no) if no is not None else "")
                name_var = tk.StringVar(value=name if name is not None else "")
                abbrev_var = tk.StringVar(value=abbrev if abbrev is not None else "")
                install_hours_var = tk.StringVar(value=str(install_hours) if install_hours is not None else "")
                labor_cost_var = tk.StringVar(value=str(labor_cost) if labor_cost is not None else "")
                misc_material_var = tk.StringVar(value=str(misc_material) if misc_material is not None else "")
                new_equipment_cost_var = tk.StringVar(
                    value=str(new_equipment_cost) if new_equipment_cost is not None else "")
                area_ratio_var = tk.StringVar(value=str(area_ratio) if area_ratio is not None else "")
                self.furniture1_unit_data.append(
                    (no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var,
                     new_equipment_cost_var, area_ratio_var))
        else:
            for data in initial_furniture1_data:
                no, name, abbrev, install_hours, labor_cost, misc_material, new_equipment_cost, area_ratio = data
                no_var = tk.StringVar(value=str(no) if no is not None else "")
                name_var = tk.StringVar(value=name if name is not None else "")
                abbrev_var = tk.StringVar(value=abbrev if abbrev is not None else "")
                install_hours_var = tk.StringVar(value=str(install_hours) if install_hours is not None else "")
                labor_cost_var = tk.StringVar(value=str(labor_cost) if labor_cost is not None else "")
                misc_material_var = tk.StringVar(value=str(misc_material) if misc_material is not None else "")
                new_equipment_cost_var = tk.StringVar(
                    value=str(new_equipment_cost) if new_equipment_cost is not None else "")
                area_ratio_var = tk.StringVar(value=str(area_ratio) if area_ratio is not None else "")
                self.furniture1_unit_data.append(
                    (no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var,
                     new_equipment_cost_var, area_ratio_var))

        # 家具設定単価2のデータ
        self.furniture2_data_file = "furniture2_data.json"
        initial_furniture2_data = [
            (1, "テーブル（自立固定）  Ｗ1600", "2人テーブル", 0.1, 27000, 2000, 19000, 0.02),
            (2, "テーブル（自立固定）  Ｗ1600", "4人テーブル", 0.1, 27000, 2000, 38000, 0.02),
            (3, "テーブル（自立固定）  Ｗ1800", "6人テーブル", 0.12, 27000, 2000, 57000, 0.02),
            (4, "ゼノンＢ（張りぐるみ）", "椅子Ａ", 0.3, 27000, 2000, 120000, 0.02),
            (5, "マニコＢ（背張）", "椅子Ｂ", 0.3, 27000, 2000, 120000, 0.02),
            (6, "椅子Ｃ", "椅子Ｃ", 0.3, 27000, 2000, 120000, 0.02),
            (7, "椅子Ｄ", "椅子Ｄ", 0.3, 27000, 2000, 120000, 0.02),
            (8, "ベンチソファＷ4060", "ソファーＡ", 0.3, 27000, 2000, 230000, 0.02),
            (9, "ベンチソファＷ3040", "ソファーＢ", 0.3, 27000, 2000, 173000, 0.02),
            (10, "ベンチソファＷ2110", "ソファーＣ", 0.3, 27000, 2000, 123000, 0.02),
            (11, "ラウンジソファ", "ソファーＤ", 0.3, 27000, 2000, 81200, 0.02),
            (12, "アペーゴサイドテーブル  ＷＣ", "ローテーブルＡ", 0.3, 27000, 2000, 33300, 0.02),
            (13, "アビタスタイルローテーブルＷ1200", "ローテーブルＢ", 0.3, 27000, 2000, 52200, 0.02),
            (14, "屋外ローテーブル", "ローテーブルＣ", 0.3, 27000, 2000, 40800, 0.02),
            (15, "コルヌラウンジチェア", "ラウンジチェアＡ", 0.3, 27000, 2000, 72900, 0.02),
            (16, "ストートラウンジチェア", "ラウンジチェアＢ", 0.3, 27000, 2000, 114800, 0.02),
            (17, "ハイバック（Ｌ字Ｗ6100+2715）", "ソファーＤ", 2, 27000, 2000, 640000, 0.02),
            (18, "ＴＶ台Ａ", "ＴＶ台Ａ", 0.3, 27000, 2000, 99000, 0.02),
            (19, "ＴＶ台Ｂ", "ＴＶ台Ｂ", 0.3, 27000, 2000, 126000, 0.02),
            (20, "ＴＶ台Ｃ", "ＴＶ台Ｃ", 0.3, 27000, 2000, 156000, 0.02),
            (21, "ヘッドボード", "ヘッドボード", 0.3, 27000, 2000, 200000, 0.02),

        ]

        self.initial_furniture2_data = initial_furniture2_data
        self.furniture2_unit_data = []

        loaded_furniture2 = self.load_furniture2_data()
        if loaded_furniture2:
            for data in loaded_furniture2:
                no, name, abbrev, install_hours, labor_cost, misc_material, new_equipment_cost, room_ratio = data
                no_var = tk.StringVar(value=str(no) if no is not None else "")
                name_var = tk.StringVar(value=name if name is not None else "")
                abbrev_var = tk.StringVar(value=abbrev if abbrev is not None else "")
                install_hours_var = tk.StringVar(value=str(install_hours) if install_hours is not None else "")
                labor_cost_var = tk.StringVar(value=str(labor_cost) if labor_cost is not None else "")
                misc_material_var = tk.StringVar(value=str(misc_material) if misc_material is not None else "")
                new_equipment_cost_var = tk.StringVar(
                    value=str(new_equipment_cost) if new_equipment_cost is not None else "")
                room_ratio_var = tk.StringVar(value=str(room_ratio) if room_ratio is not None else "")
                self.furniture2_unit_data.append(
                    (no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var,
                     new_equipment_cost_var, room_ratio_var))
        else:
            for data in initial_furniture2_data:
                no, name, abbrev, install_hours, labor_cost, misc_material, new_equipment_cost, room_ratio = data
                no_var = tk.StringVar(value=str(no) if no is not None else "")
                name_var = tk.StringVar(value=name if name is not None else "")
                abbrev_var = tk.StringVar(value=abbrev if abbrev is not None else "")
                install_hours_var = tk.StringVar(value=str(install_hours) if install_hours is not None else "")
                labor_cost_var = tk.StringVar(value=str(labor_cost) if labor_cost is not None else "")
                misc_material_var = tk.StringVar(value=str(misc_material) if misc_material is not None else "")
                new_equipment_cost_var = tk.StringVar(
                    value=str(new_equipment_cost) if new_equipment_cost is not None else "")
                room_ratio_var = tk.StringVar(value=str(room_ratio) if room_ratio is not None else "")
                self.furniture2_unit_data.append(
                    (no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var,
                     new_equipment_cost_var, room_ratio_var))


        # 設備機器着脱設定単価データ(初期値)
        initial_equipment_data = [
            (1, "洋風便器洗浄便座一体型", "便器", 1.56, 27000, 4212, 120000, 0.02000),
            (2, "洗面器(平付・バック無)大形 自動水栓", "洗面器", 0.69, 27000, 1863, 50000, 0.20000),
            (3, "電気温水器 10L", "温水器", 0.45, 27000, 1215, 148800, None),
            (4, "天カセAC 2.8Kw以下", "個別エアコン", 0.86, 27000, 2322, 150000, 0.05),
            (5, "天井FCU", "FCU", 1.19, 27000, 3213, 59000, 0.01786),
            (6, "天井換気扇", "換気扇", 0.50, 27000, 1350, 35000, 0.08000),
            (7, "スプリンクラーヘッド", "スプリンクラー", 0.08, 27000, 216, 8000, 0.14286),
            (8, "屋内消火栓箱", "消火栓箱", 1.40, 27000, 3780, 199800, 0.01667),
            (9, "天カセAC 5.6Kw以下", "共用部エアコン", 1.30, 27000, 3510, 300000, 0.01786),
            (10, "制気口", "制気口", 0.20, 27000, 540, 10000, 0.08000),
            (11, "排煙口", "排煙口", 0.30, 27000, 810, 20000, 0.00400),
            (12, "シャワー水栓", "シャワー水栓", 1.00, 27000, 2700, 41480, 0.14000),
            (13, "隔て板", "隔て板", 1, 27000, 8000, 200000, None),
            (14, "主室照明", "主室照明", 0.15, 27000, 405, 27450, 0.10000),
            (15, "洗面所照明", "洗面所照明", 0.15, 27000, 405, 13224, 0.20000),
            (16, "コンセント", "コンセント", 0.08, 27000, 216, 1000, 0.02000),
            (17, "スイッチ", "スイッチ", 0.08, 27000, 216, 1000, 0.02000),
            (18, "煙感知器", "煙感知器", 0.16, 27000, 429, 12915, 0.02500),
            (19, "誘導灯", "誘導灯", 0.16, 27000, 429, 12915, 0.02500),
            (20, "スピーカー", "スピーカー", 0.16, 27000, 429, 12915, 0.02500),
            (21, "熱感知器", "熱感知器", 0.13, 27000, 359, 9945, 0.02500),
            (22, "洗い場照明", "洗い場照明", 0.25, 27000, 800, 5000, None),
            (23, "通路部屋入口照明", "和ブラケット", 0.15, 27000, 405, 8460, 0.10000),
            (24, "通路天井照明", "通路LED", 0.27, 27000, 729, 13224, 0.10000),
            (25, "エントランス照明", "エントランスLED", 0.27, 27000, 729, 13224, 0.10000),
            (26, "レストラン照明", "レストランLED", 0.27, 27000, 729, 13224, 0.10000),
            (27, "厨房照明", "厨房照明", 0.27, 27000, 729, 9000, 0.20000),
            (28, "厨房フード", "フード", 2.0, 27000, 30000, 300000, 0.02),
        ]

        self.initial_equipment_data = initial_equipment_data
        self.equipment_unit_data = []
        self.equipment_data_file = "equipment_data.json"

        loaded_data = self.load_equipment_data()

        if loaded_data:
            for data in loaded_data:
                no, name, abbrev, install_hours, labor_cost, misc_material, new_equipment_cost, area_ratio = data
                no_var = tk.StringVar(value=str(no) if no is not None else "")
                name_var = tk.StringVar(value=name if name is not None else "")
                abbrev_var = tk.StringVar(value=abbrev if abbrev is not None else "")
                install_hours_var = tk.StringVar(value=str(install_hours) if install_hours is not None else "")
                labor_cost_var = tk.StringVar(value=str(labor_cost) if labor_cost is not None else "")
                misc_material_var = tk.StringVar(value=str(misc_material) if misc_material is not None else "")
                new_equipment_cost_var = tk.StringVar(
                    value=str(new_equipment_cost) if new_equipment_cost is not None else "")
                area_ratio_var = tk.StringVar(value=str(area_ratio) if area_ratio is not None else "")

                self.equipment_unit_data.append((
                    no_var, name_var, abbrev_var, install_hours_var,
                    labor_cost_var, misc_material_var, new_equipment_cost_var, area_ratio_var
                ))
        else:
            for data in initial_equipment_data:
                no, name, abbrev, install_hours, labor_cost, misc_material, new_equipment_cost, area_ratio = data

                no_var = tk.StringVar(value=str(no) if no is not None else "")
                name_var = tk.StringVar(value=name if name is not None else "")
                abbrev_var = tk.StringVar(value=abbrev if abbrev is not None else "")
                install_hours_var = tk.StringVar(value=str(install_hours) if install_hours is not None else "")
                labor_cost_var = tk.StringVar(value=str(labor_cost) if labor_cost is not None else "")
                misc_material_var = tk.StringVar(value=str(misc_material) if misc_material is not None else "")
                new_equipment_cost_var = tk.StringVar(
                    value=str(new_equipment_cost) if new_equipment_cost is not None else "")
                area_ratio_var = tk.StringVar(value=str(area_ratio) if area_ratio is not None else "")

                self.equipment_unit_data.append((
                    no_var, name_var, abbrev_var, install_hours_var,
                    labor_cost_var, misc_material_var, new_equipment_cost_var, area_ratio_var
                ))



        # 熱源機器設定のデータ
        self.heat_source_data_file = "heat_source_data.json"
        initial_heat_source_data = [

            (1, "CVM-3003M_A", "真空ヒーター349Kw", 7.80, 27000, 2000, 3890000, 0.02, "重油", "349Kw"),
            (2, "CVM-4003M_A", "真空ヒーター465Kw", 9.00, 27000, 2000, 5727000, 0.02, "重油", "465Kw"),
            (3, "CVM-5003M_A", "真空ヒーター581Kw", 10.20, 27000, 2000, 6315000, 0.02, "重油", "581Kw"),
            (4, "CVM-3003M_k", "真空ヒーター349Kw", 7.80, 27000, 2000, 3890000, 0.02, "灯油", "349Kw"),
            (5, "CVM-4003M_k", "真空ヒーター465Kw", 9.00, 27000, 2000, 5727000, 0.02, "灯油", "465Kw"),
            (6, "CVM-5003M_k", "真空ヒーター581Kw", 10.20, 27000, 2000, 6315000, 0.02, "灯油", "581Kw"),
            (7, "CVM-3003M_G", "真空ヒーター349Kw", 7.80, 27000, 2000, 3890000, 0.02, "ガス", "349Kw"),
            (8, "CVM-4003M_G", "真空ヒーター465Kw", 9.00, 27000, 2000, 5727000, 0.02, "ガス", "465Kw"),
            (9, "CVM-5003M_G", "真空ヒーター581Kw", 10.20, 27000, 2000, 6315000, 0.02, "ガス", "581Kw"),
            (10, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02, "", ""),
            (11, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02, "", ""),
            (12, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02, "", ""),
            (13, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02, "", ""),
            (14, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02, "", ""),
            (15, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02, "", ""),
            (16, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02, "", ""),
            (17, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02, "", ""),
            (18, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02, "", ""),
            (19, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02, "", ""),
            (20, "〇〇〇", "〇〇〇", 0.1, 27000, 2000, 30000, 0.02, "", ""),
        ]

        self.initial_heat_source_data = initial_heat_source_data
        self.heat_source_unit_data = []

        loaded_heat_source = self.load_heat_source_data()
        if loaded_heat_source:
            for data in loaded_heat_source:
                no, name, abbrev, install_hours, labor_cost, misc_material, new_equipment_cost, area_ratio, fuel_type, output = data  # ← output追加
                no_var = tk.StringVar(value=str(no) if no is not None else "")
                name_var = tk.StringVar(value=name if name is not None else "")
                abbrev_var = tk.StringVar(value=abbrev if abbrev is not None else "")
                install_hours_var = tk.StringVar(value=str(install_hours) if install_hours is not None else "")
                labor_cost_var = tk.StringVar(value=str(labor_cost) if labor_cost is not None else "")
                misc_material_var = tk.StringVar(value=str(misc_material) if misc_material is not None else "")
                new_equipment_cost_var = tk.StringVar(
                    value=str(new_equipment_cost) if new_equipment_cost is not None else "")
                area_ratio_var = tk.StringVar(value=str(area_ratio) if area_ratio is not None else "")
                fuel_type_var = tk.StringVar(value=str(fuel_type) if fuel_type is not None else "")
                output_var = tk.StringVar(value=str(output) if output is not None else "")  # ← 追加
                self.heat_source_unit_data.append(
                    (no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var,
                     new_equipment_cost_var, area_ratio_var, fuel_type_var, output_var))  # ← output_var追加
        else:
            for data in initial_heat_source_data:
                no, name, abbrev, install_hours, labor_cost, misc_material, new_equipment_cost, area_ratio, fuel_type, output = data  # ← output追加
                no_var = tk.StringVar(value=str(no) if no is not None else "")
                name_var = tk.StringVar(value=name if name is not None else "")
                abbrev_var = tk.StringVar(value=abbrev if abbrev is not None else "")
                install_hours_var = tk.StringVar(value=str(install_hours) if install_hours is not None else "")
                labor_cost_var = tk.StringVar(value=str(labor_cost) if labor_cost is not None else "")
                misc_material_var = tk.StringVar(value=str(misc_material) if misc_material is not None else "")
                new_equipment_cost_var = tk.StringVar(
                    value=str(new_equipment_cost) if new_equipment_cost is not None else "")
                area_ratio_var = tk.StringVar(value=str(area_ratio) if area_ratio is not None else "")
                fuel_type_var = tk.StringVar(value=str(fuel_type) if fuel_type is not None else "")
                output_var = tk.StringVar(value=str(output) if output is not None else "")  # ← 追加
                self.heat_source_unit_data.append(
                    (no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var,
                     new_equipment_cost_var, area_ratio_var, fuel_type_var, output_var))  # ← output_var追加

        # 厨房機器設定のデータ
        self.kitchen_equipment_data_file = "kitchen_equipment_data.json"
        initial_kitchen_equipment_data = [
            # 新設機器単価の列を削除し、8番目の要素を削除
            (1, "コールドテーブル", "コールドテーブル", 0.74, 27000, 36000, 600000, 0.60, "", ""),
            (2, "業務用冷蔵庫W1800", "業務用冷蔵庫W1800", 1.48, 27000, 120000, 2000000, 0.60, "", ""),
            (3, "業務用冷凍庫W1200", "業務用冷凍庫W1200", 1.48, 27000, 120000, 2000000, 0.60, "", ""),
            (4, "冷蔵ショーケース1100", "冷蔵ショーケース1100", 1.48, 27000, 60000, 1000000, 0.60, "", ""),
            (5, "ベーカリー", "ベーカリー", 1.48, 27000, 120000, 2000000, 0.60, "18.0 Kw", "電気"),
            (6, "ピザオーブン", "ピザオーブン", 1.48, 27000, 120000, 2000000, 0.60, "18.0 Kw", "電気"),
            (7, "電気スチコン", "電気スチコン", 1.48, 27000, 120000, 2000000, 0.60, "12.0 Kw", "電気"),
            (8, "寿司メーカー", "寿司メーカー", 1.11, 27000, 60000, 1000000, 0.60, "", ""),
            (9, "ガスフライヤー", "ガスフライヤー", 0.74, 27000, 12000, 200000, 0.60, "6.4 Kw", "ガス"),
            (10, "ガスレンジ", "ガスレンジ", 0.74, 27000, 12000, 200000, 0.60, "28.0 Kw", "ガス"),
            (11, "スープレンジ", "スープレンジ", 0.74, 27000, 6000, 100000, 0.60, "18.0 Kw", "ガス"),
            (12, "茹で麺機", "茹で麺機", 0.74, 27000, 30000, 500000, 0.60, "28.0 Kw", "ガス"),
            (13, "焼台", "焼台", 0.74, 27000, 12000, 200000, 0.60, "4.0 Kw", "ガス"),
            (14, "炊飯器", "炊飯器", 0.74, 27000, 9000, 150000, 0.60, "11.0 Kw", "ガス"),
            (15, "食器洗浄機", "食器洗浄機", 8.00, 27000, 375001, 5000000, 0.75, "50.0 Kw", "ガス"),
            (16, "コンロ台900", "コンロ台900", 0.30, 27000, 4500, 60000, 0.75, "", ""),
            (17, "シェルフ1800×600", "シェルフ1800×600", 0.30, 27000, 3750, 50000, 0.75, "", ""),
            (18, "作業台1200", "作業台1200", 0.74, 27000, 5250, 70000, 0.75, "", ""),
            (19, "1槽シンク", "1槽シンク", 0.37, 27000, 7500, 100000, 0.75, "", ""),
            (20, "3槽シンク", "3槽シンク", 0.44, 27000, 9000, 120000, 0.75, "", ""),
            (21, "吊戸棚", "吊戸棚", 0.74, 27000, 5775, 77000, 0.75, "", ""),
            (22, "パイプ棚", "パイプ棚", 0.37, 27000, 1500, 20000, 0.75, "", ""),
            (23, "プレハブ冷蔵庫(3坪)", "プレハブ冷蔵庫(3坪)", 4.00, 27000, 180000, 3000000, 0.60, "", ""),
        ]

        self.initial_kitchen_equipment_data = initial_kitchen_equipment_data
        self.kitchen_equipment_unit_data = []
        self.kitchen_cost_display_vars = []

        loaded_kitchen_equipment = self.load_kitchen_equipment_data()
        if loaded_kitchen_equipment:
            for data in loaded_kitchen_equipment:
                no, name, abbrev, install_hours, labor_cost, expense, list_price, rate, unit_input, category = data
                no_var = tk.StringVar(value=str(no) if no is not None else "")
                name_var = tk.StringVar(value=name if name is not None else "")
                abbrev_var = tk.StringVar(value=abbrev if abbrev is not None else "")
                install_hours_var = tk.StringVar(value=str(install_hours) if install_hours is not None else "")
                labor_cost_var = tk.StringVar(value=str(labor_cost) if labor_cost is not None else "")
                expense_var = tk.StringVar(value=str(expense) if expense is not None else "")
                list_price_var = tk.StringVar(value=str(list_price) if list_price is not None else "")
                rate_var = tk.StringVar(value=str(rate) if rate is not None else "")
                unit_input_var = tk.StringVar(value=str(unit_input) if unit_input is not None else "")
                category_var = tk.StringVar(value=str(category) if category is not None else "")
                self.kitchen_equipment_unit_data.append(
                    (no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, expense_var,
                     list_price_var, rate_var, unit_input_var, category_var))
        else:
            for data in initial_kitchen_equipment_data:
                no, name, abbrev, install_hours, labor_cost, expense, list_price, rate, unit_input, category = data
                no_var = tk.StringVar(value=str(no) if no is not None else "")
                name_var = tk.StringVar(value=name if name is not None else "")
                abbrev_var = tk.StringVar(value=abbrev if abbrev is not None else "")
                install_hours_var = tk.StringVar(value=str(install_hours) if install_hours is not None else "")
                labor_cost_var = tk.StringVar(value=str(labor_cost) if labor_cost is not None else "")
                expense_var = tk.StringVar(value=str(expense) if expense is not None else "")
                list_price_var = tk.StringVar(value=str(list_price) if list_price is not None else "")
                rate_var = tk.StringVar(value=str(rate) if rate is not None else "")
                unit_input_var = tk.StringVar(value=str(unit_input) if unit_input is not None else "")
                category_var = tk.StringVar(value=str(category) if category is not None else "")
                self.kitchen_equipment_unit_data.append(
                    (no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, expense_var,
                     list_price_var, rate_var, unit_input_var, category_var))

        # 【追加】建築構造別単価データ定義
        # 行: 木造、RC造、鉄骨造、SRC造
        # 列: Gensen, Premium1, Premium2, TAOYA1, TAOYA2 (円/m2)
        initial_structure_unit_cost_data = [
            ("木造", 300000, 350000, 400000, 450000, 500000),
            ("RC造", 500000, 550000, 600000, 650000, 700000),
            ("鉄骨造", 550000, 600000, 650000, 700000, 750000),
            ("SRC造", 600000, 650000, 700000, 750000, 800000)
        ]

        self.structure_unit_costs = []
        for structure, gensen, p1, p2, t1, t2 in initial_structure_unit_cost_data:
            # 5つのグレードごとの単価を格納するStringVarリストを作成
            cost_vars = [
                tk.StringVar(value=str(gensen)),
                tk.StringVar(value=str(p1)),
                tk.StringVar(value=str(p2)),
                tk.StringVar(value=str(t1)),
                tk.StringVar(value=str(t2)),
            ]
            self.structure_unit_costs.append((structure, cost_vars))


        # 建築内装単価データ(初期値)
        initial_unit_cost_data = [
            ("通路", 2.5, 1800, 2060, 2060, 3100, 3100, 3100, 3490, 3490, 3750, 3750, 3100, 3490, 3490, 3750, 3750),
            ("客室", 2.5, 4400, 5050, 5050, 6350, 6350, 4400, 5050, 5050, 7000, 7000, 4400, 5050, 5050, 7000, 7000),
            ("アミューズメント", 2.5, 1800, 2060, 2060, 3100, 3100, 3100, 3490, 3490, 3750, 3750, 3100, 3490, 3490,
             3750, 3750),
            ("レセプション", 3.5, 1800, 2060, 2060, 3100, 3100, 3100, 3490, 3490, 3750, 3750, 3100, 3490, 3490, 3750,
             3750),
            ("レストラン", 3.5, 1800, 2060, 2060, 3100, 3100, 3100, 3490, 3490, 3750, 3750, 3100, 3490, 3490, 3750,
             3750),
            ("事務所", 2.5, 1800, 1800, 1800, 1800, 1800, 3100, 3100, 3100, 3100, 3100, 3100, 3100, 3100, 3100, 3100),
            ("会議室", 2.7, 1800, 2060, 2060, 3100, 3100, 3100, 3490, 3490, 3750, 3750, 3100, 3490, 3490, 3750, 3750),
            ("浴室", 4.5, 12000, 18000, 18000, 18000, 18000, 25000, 25000, 25000, 25000, 25000, 35000, 35000, 35000,
             40000, 40000),
            ("脱衣室", 2.5, 1800, 2060, 2060, 3100, 3100, 3100, 3490, 3490, 3750, 3750, 3100, 3490, 3490, 3750, 3750),
            ("湯上処", 2.5, 1800, 2060, 2060, 3100, 3100, 3100, 3490, 3490, 3750, 3750, 3100, 3490, 3490, 3750, 3750),
            ("露天風呂", 4.5, 12000, 18000, 18000, 18000, 18000, 25000, 25000, 25000, 25000, 25000, 35000, 35000, 35000,
             40000, 40000),
            ("内風呂", 2.5, 4400, 5050, 5050, 6350, 6350, 4400, 5050, 5050, 7000, 7000, 4400, 5050, 5050, 7000, 7000),
            ("岩盤浴", 2.5, 4400, 5050, 5050, 6350, 6350, 4400, 5050, 5050, 7000, 7000, 4400, 5050, 5050, 7000, 7000),
            ("サウナ", 2.5, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000),
            ("ラウンジ", 3.5, 1800, 2060, 2060, 3100, 3100, 3100, 3490, 3490, 3750, 3750, 3100, 3490, 3490, 3750, 3750),
            ("厨房", 2.5, 1800, 1800, 1800, 1800, 1800, 3100, 3100, 3100, 3100, 3100, 3100, 3100, 3100, 3100, 3100),
            ("売店", 2.5, 1800, 2060, 2060, 3100, 3100, 3100, 3490, 3490, 3750, 3750, 3100, 3490, 3490, 3750, 3750),
            ("フロント", 2.5, 1800, 2060, 2060, 3100, 3100, 3100, 3490, 3490, 3750, 3750, 3100, 3490, 3490, 3750, 3750),
            ("ロビー", 2.5, 1800, 2060, 2060, 3100, 3100, 3100, 3490, 3490, 3750, 3750, 3100, 3490, 3490, 3750, 3750),
            ("ファサード", 2.5, 1800, 2060, 2060, 3100, 3100, 3100, 3490, 3490, 3750, 3750, 3100, 3490, 3490, 3750,
             3750),
            ("ライブキッチン", 2.5, 1800, 2060, 2060, 3100, 3100, 3100, 3490, 3490, 3750, 3750, 3100, 3490, 3490, 3750,
             3750),
            ("和室(10-15J)", 2.5, 4400, 5050, 5050, 6350, 6350, 4400, 5050, 5050, 7000, 7000, 4400, 5050, 5050, 7000,
             7000),
            ("和室・洋室", 2.5, 4400, 5050, 5050, 6350, 6350, 4400, 5050, 5050, 7000, 7000, 4400, 5050, 5050, 7000,
             7000),
            ("和ベッド", 2.5, 4400, 5050, 5050, 6350, 6350, 4400, 5050, 5050, 7000, 7000, 4400, 5050, 5050, 7000, 7000),
            ("洋室", 2.5, 4400, 5050, 5050, 6350, 6350, 4400, 5050, 5050, 7000, 7000, 4400, 5050, 5050, 7000, 7000),
            ("卓球コーナー", 2.5, 1800, 2060, 2060, 3100, 3100, 3100, 3490, 3490, 3750, 3750, 3100, 3490, 3490,
             3750, 3750),
            ("キッズスペース", 2.5, 1800, 2060, 2060, 3100, 3100, 3100, 3490, 3490, 3750, 3750, 3100, 3490, 3490,
             3750, 3750),
            ("漫画コーナー", 2.5, 1800, 2060, 2060, 3100, 3100, 3100, 3490, 3490, 3750, 3750, 3100, 3490, 3490,
             3750, 3750),

        ]

        self.initial_unit_cost_data = initial_unit_cost_data
        self.arch_unit_costs = []
        self.arch_data_file = "arch_unit_cost_data.json"

        loaded_arch_data = self.load_arch_data()

        if loaded_arch_data:
            for data in loaded_arch_data:
                name, height, *costs = data
                height_var = tk.StringVar(value=str(height))
                cost_vars = [tk.StringVar(value=str(c)) for c in costs]
                self.arch_unit_costs.append((name, height_var, cost_vars))
        else:
            for name, height, *costs in initial_unit_cost_data:
                height_var = tk.StringVar(value=str(height))
                cost_vars = [tk.StringVar(value=str(c)) for c in costs]
                self.arch_unit_costs.append((name, height_var, cost_vars))

        self.arch_base_cost_data = [
            "LGS天井(PB9.5)", "LGS天井(PB12.5)", "LGS壁(PB12.5+9.5)", "戸境壁",
            "レストラン", "事務所", "会議室", "浴室", "脱衣室", "湯上処",
            "露天風呂", "内風呂", "岩盤浴", "サウナ", "ラウンジ", "厨房",
            "売店", "ファサード", "ライブキッチン", "和室(10-15J)", "和室・洋室",
            "和ベッド", "洋室"
        ]


        # ノートブック(タブ)作成
        self.notebook = ttk.Notebook(main_container)  # ← rootをmain_containerに変更
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # ★★★ タブのスタイル設定（境界線を完全に削除）★★★
        style = ttk.Style()

        # 統一する背景色を定義
        unified_bg = '#f0f0f0'

        # ★★★ タブバー全体の背景（境界線を削除）★★★
        style.configure('TNotebook',
                        background=unified_bg,
                        borderwidth=0,  # 外側の境界線なし
                        relief='flat',  # フラット表示
                        tabmargins=[0, 0, 0, 0],  # タブ間のマージンを0に
                        highlightthickness=0)  # ハイライト境界線なし

        # ★★★ タブ自体のスタイル（非選択時）- 境界線を完全に削除 ★★★
        style.configure('TNotebook.Tab',
                        background=unified_bg,
                        foreground='#999999',
                        padding=[0, 0],
                        borderwidth=0,  # タブの境界線なし
                        relief='flat',  # フラット表示
                        focuscolor='',  # フォーカス時の枠線なし
                        highlightthickness=0)  # ハイライト境界線なし

        # ★★★ タブのスタイル（選択時）- 境界線なし ★★★
        style.map('TNotebook.Tab',
                  background=[('selected', unified_bg)],
                  foreground=[('selected', '#000000')],
                  relief=[('selected', 'flat')],
                  borderwidth=[('selected', 0)])  # 選択時も境界線なし

        # コンテンツエリアの背景色を統一
        style.configure('TFrame', background=unified_bg)
        # ★★★ ここまで修正 ★★★

        # ★ 物件管理タブを最初に追加 ★
        self.property_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.property_frame, text='物件管理')

        self.calc_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.calc_frame, text='コストサマリ')

        self.onsen_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.onsen_frame, text='温泉設備')

        self.restaurant_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.restaurant_frame, text='レストラン')

        self.lounge_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.lounge_frame, text='ラウンジ')

        self.hallway_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.hallway_frame, text='通路')

        self.amusement_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.amusement_frame, text='アミューズメント')

        self.guest_room_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.guest_room_frame, text='客室')

        self.arch_settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.arch_settings_frame, text='建築初期設定')

        self.equipment_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.equipment_frame, text='設備設定')

        self.furniture_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.furniture_frame, text='家具')

        self.settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text='経費設定')

        self.lcc_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.lcc_frame, text='LCC分析')


        # 温泉設備タブの計算値表示用変数
        self.men_calculated_bath_capacity = tk.StringVar(value="計算値: 0 L")
        self.men_calculated_outdoor_capacity = tk.StringVar(value="計算値: 0 L")
        self.men_calculated_bathroom_area = tk.StringVar(value="計算値: 0 ㎡")
        self.women_calculated_bath_capacity = tk.StringVar(value="計算値: 0 L")
        self.women_calculated_outdoor_capacity = tk.StringVar(value="計算値: 0 L")
        self.women_calculated_bathroom_area = tk.StringVar(value="計算値: 0 ㎡")

        self.men_recommended_indoor_capacity = tk.StringVar(value="推奨値: 0 L")
        self.men_recommended_indoor_tub_area = tk.StringVar(value="推奨値: 0 ㎡")#男内湯の推奨面積
        self.men_recommended_outdoor_capacity = tk.StringVar(value="推奨値: 0 L")
        self.men_recommended_outdoor_tub_area = tk.StringVar(value="推奨値: 0 ㎡")  # 男外湯の推奨面積

        self.women_recommended_indoor_capacity = tk.StringVar(value="推奨値: 0 L")
        self.women_recommended_indoor_tub_area = tk.StringVar(value="推奨値: 0 ㎡")#女内湯の推奨面積
        self.women_recommended_outdoor_capacity = tk.StringVar(value="推奨値: 0 L")
        self.women_recommended_outdoor_tub_area = tk.StringVar(value="推奨値: 0 ㎡")  # 女外湯の推奨面積

        self.men_recommended_bathroom_area = tk.StringVar(value="推奨値: 0 ㎡")
        self.women_recommended_bathroom_area = tk.StringVar(value="推奨値: 0 ㎡")

        self.men_calculation_process = tk.StringVar(value="")
        self.women_calculation_process = tk.StringVar(value="")

        self.create_property_tab()# 物件管理タブの作成
        self.create_calculation_tab()#コストサマリタブの作成
        self.create_arch_settings_tab()
        self.create_settings_tab()
        self.create_onsen_tab()
        self.create_restaurant_tab()
        self.create_lounge_tab()
        self.create_hallway_tab()
        self.create_amusement_tab()
        self.create_guest_room_tab()
        self.create_equipment_tab()
        self.create_furniture_tab()
        self.create_lcc_tab()

        self.update_onsen_calculations()

        self.capacity_japanese_room.trace('w', self.update_guest_room_capacities)
        self.capacity_japanese_western_room.trace('w', self.update_guest_room_capacities)
        self.capacity_japanese_bed_room.trace('w', self.update_guest_room_capacities)
        self.capacity_western_room.trace('w', self.update_guest_room_capacities)
        self.update_guest_room_capacities()

        # 温泉設備のトレース設定
        self.men_partition_count.trace('w', self.update_wash_equipment_costs)
        self.men_wash_lighting_count.trace('w', self.update_wash_equipment_costs)
        self.men_shower_faucet_count.trace('w', self.update_wash_equipment_costs)
        self.women_partition_count.trace('w', self.update_wash_equipment_costs)
        self.women_wash_lighting_count.trace('w', self.update_wash_equipment_costs)
        self.women_shower_faucet_count.trace('w', self.update_wash_equipment_costs)
        # 特殊設備のチェックボックス変更時に金額計算と小計更新
        self.men_sauna.trace('w', self.update_special_equipment_costs)
        self.men_jacuzzi.trace('w', self.update_special_equipment_costs)
        self.men_water_bath.trace('w', self.update_special_equipment_costs)
        self.women_sauna.trace('w', self.update_special_equipment_costs)
        self.women_jacuzzi.trace('w', self.update_special_equipment_costs)
        self.women_water_bath.trace('w', self.update_special_equipment_costs)

        # 特殊設備の金額が更新されたときに小計を更新
        self.men_sauna_cost.trace('w', self.update_onsen_subtotal)
        self.men_jacuzzi_cost.trace('w', self.update_onsen_subtotal)
        self.men_water_bath_cost.trace('w', self.update_onsen_subtotal)
        self.women_sauna_cost.trace('w', self.update_onsen_subtotal)
        self.women_jacuzzi_cost.trace('w', self.update_onsen_subtotal)
        self.women_water_bath_cost.trace('w', self.update_onsen_subtotal)

        # ラウンジタブのトレース設定
        self.lounge_arch_front_area.trace('w', self.update_lounge_arch_costs)
        self.lounge_arch_lobby_area.trace('w', self.update_lounge_arch_costs)
        self.lounge_arch_facade_area.trace('w', self.update_lounge_arch_costs)
        self.lounge_arch_shop_area.trace('w', self.update_lounge_arch_costs)

        self.lounge_mech_ac_count.trace('w', self.update_lounge_equipment_costs)
        self.lounge_mech_sprinkler_count.trace('w', self.update_lounge_equipment_costs)
        self.lounge_mech_fire_hose_count.trace('w', self.update_lounge_equipment_costs)
        self.lounge_elec_led_count.trace('w', self.update_lounge_equipment_costs)
        self.lounge_elec_smoke_count.trace('w', self.update_lounge_equipment_costs)
        self.lounge_elec_exit_light_count.trace('w', self.update_lounge_equipment_costs)
        self.lounge_elec_speaker_count.trace('w', self.update_lounge_equipment_costs)
        self.lounge_elec_outlet_count.trace('w', self.update_lounge_equipment_costs)

        self.lounge_arch_front_check.trace('w', self.update_lounge_subtotal)
        self.lounge_arch_lobby_check.trace('w', self.update_lounge_subtotal)
        self.lounge_arch_facade_check.trace('w', self.update_lounge_subtotal)
        self.lounge_arch_shop_check.trace('w', self.update_lounge_subtotal)
        self.lounge_mech_ac_check.trace('w', self.update_lounge_subtotal)
        self.lounge_mech_sprinkler_check.trace('w', self.update_lounge_subtotal)
        self.lounge_mech_fire_hose_check.trace('w', self.update_lounge_subtotal)
        self.lounge_elec_led_check.trace('w', self.update_lounge_subtotal)
        self.lounge_elec_smoke_check.trace('w', self.update_lounge_subtotal)
        self.lounge_elec_exit_light_check.trace('w', self.update_lounge_subtotal)
        self.lounge_elec_speaker_check.trace('w', self.update_lounge_subtotal)
        self.lounge_elec_outlet_check.trace('w', self.update_lounge_subtotal)

        self.update_lounge_arch_costs()
        self.update_lounge_recommended_counts()
        self.update_lounge_equipment_costs()

        # 通路タブのトレース設定
        self.hallway_mech_ac_count.trace('w', self.update_hallway_equipment_costs)
        self.hallway_mech_sprinkler_count.trace('w', self.update_hallway_equipment_costs)
        self.hallway_mech_fire_hose_count.trace('w', self.update_hallway_equipment_costs)
        self.hallway_elec_led_count.trace('w', self.update_hallway_equipment_costs)
        self.hallway_elec_smoke_count.trace('w', self.update_hallway_equipment_costs)
        self.hallway_elec_exit_light_count.trace('w', self.update_hallway_equipment_costs)
        self.hallway_elec_speaker_count.trace('w', self.update_hallway_equipment_costs)
        self.hallway_elec_outlet_count.trace('w', self.update_hallway_equipment_costs)

        # 通路のチェックボックス監視
        self.hallway_mech_ac_check.trace('w', self.update_hallway_subtotal)
        self.hallway_mech_sprinkler_check.trace('w', self.update_hallway_subtotal)
        self.hallway_mech_fire_hose_check.trace('w', self.update_hallway_subtotal)
        self.hallway_elec_led_check.trace('w', self.update_hallway_subtotal)
        self.hallway_elec_smoke_check.trace('w', self.update_hallway_subtotal)
        self.hallway_elec_exit_light_check.trace('w', self.update_hallway_subtotal)
        self.hallway_elec_speaker_check.trace('w', self.update_hallway_subtotal)
        self.hallway_elec_outlet_check.trace('w', self.update_hallway_subtotal)

        # 階数変更時に通路タブを更新
        self.floors.trace('w', self.update_hallway_floors)
        self.underground_floors.trace('w', self.update_hallway_floors)

        self.update_hallway_floors()
        self.update_hallway_recommended_counts()
        self.update_hallway_equipment_costs()

        # アミューズメントタブのトレース設定
        self.amusement_arch_pingpong_area.trace('w', self.update_amusement_arch_costs)
        self.amusement_arch_kids_area.trace('w', self.update_amusement_arch_costs)
        self.amusement_arch_manga_area.trace('w', self.update_amusement_arch_costs)

        self.amusement_mech_ac_count.trace('w', self.update_amusement_equipment_costs)
        self.amusement_mech_sprinkler_count.trace('w', self.update_amusement_equipment_costs)
        self.amusement_mech_fire_hose_count.trace('w', self.update_amusement_equipment_costs)
        self.amusement_elec_led_count.trace('w', self.update_amusement_equipment_costs)
        self.amusement_elec_smoke_count.trace('w', self.update_amusement_equipment_costs)
        self.amusement_elec_exit_light_count.trace('w', self.update_amusement_equipment_costs)
        self.amusement_elec_speaker_count.trace('w', self.update_amusement_equipment_costs)
        self.amusement_elec_outlet_count.trace('w', self.update_amusement_equipment_costs)

        # アミューズメントのチェックボックス監視
        self.amusement_arch_pingpong_check.trace('w', self.update_amusement_subtotal)
        self.amusement_arch_kids_check.trace('w', self.update_amusement_subtotal)
        self.amusement_arch_manga_check.trace('w', self.update_amusement_subtotal)
        self.amusement_mech_ac_check.trace('w', self.update_amusement_subtotal)
        self.amusement_mech_sprinkler_check.trace('w', self.update_amusement_subtotal)
        self.amusement_mech_fire_hose_check.trace('w', self.update_amusement_subtotal)
        self.amusement_elec_led_check.trace('w', self.update_amusement_subtotal)
        self.amusement_elec_smoke_check.trace('w', self.update_amusement_subtotal)
        self.amusement_elec_exit_light_check.trace('w', self.update_amusement_subtotal)
        self.amusement_elec_speaker_check.trace('w', self.update_amusement_subtotal)
        self.amusement_elec_outlet_check.trace('w', self.update_amusement_subtotal)

        self.update_amusement_arch_costs()
        self.update_amusement_recommended_counts()
        self.update_amusement_equipment_costs()

        # 客室タブのトレース設定
        self.guest_arch_japanese_area.trace('w', self.update_guest_arch_costs)
        self.guest_arch_japanese_western_area.trace('w', self.update_guest_arch_costs)
        self.guest_arch_japanese_bed_area.trace('w', self.update_guest_arch_costs)
        self.guest_arch_western_area.trace('w', self.update_guest_arch_costs)

        self.guest_mech_ac_count.trace('w', self.update_guest_equipment_costs)
        self.guest_mech_wash_basin_count.trace('w', self.update_guest_equipment_costs)
        self.guest_mech_sprinkler_count.trace('w', self.update_guest_equipment_costs)
        self.guest_elec_main_light_count.trace('w', self.update_guest_equipment_costs)
        self.guest_elec_smoke_count.trace('w', self.update_guest_equipment_costs)
        self.guest_elec_heat_detector_count.trace('w', self.update_guest_equipment_costs)

        self.guest_elec_speaker_count.trace('w', self.update_guest_equipment_costs)
        self.guest_elec_outlet_count.trace('w', self.update_guest_equipment_costs)

        # 客室のチェックボックス監視
        self.guest_arch_japanese_check.trace('w', self.update_guest_subtotal)
        self.guest_arch_japanese_western_check.trace('w', self.update_guest_subtotal)
        self.guest_arch_japanese_bed_check.trace('w', self.update_guest_subtotal)
        self.guest_arch_western_check.trace('w', self.update_guest_subtotal)
        self.guest_mech_ac_check.trace('w', self.update_guest_subtotal)
        self.guest_mech_wash_basin_check.trace('w', self.update_guest_subtotal)
        self.guest_mech_sprinkler_check.trace('w', self.update_guest_subtotal)
        self.guest_elec_main_light_check.trace('w', self.update_guest_subtotal)
        self.guest_elec_smoke_check.trace('w', self.update_guest_subtotal)
        self.guest_elec_heat_detector_check.trace('w', self.update_guest_subtotal)
        self.guest_elec_speaker_check.trace('w', self.update_guest_subtotal)
        self.guest_elec_outlet_check.trace('w', self.update_guest_subtotal)

        self.update_guest_arch_costs()
        self.update_guest_recommended_counts()
        self.update_guest_equipment_costs()

        # 付加設備のトレース設定
        self.japanese_room_count.trace('w', self.update_additional_equipment_costs)
        self.japanese_western_room_count.trace('w', self.update_additional_equipment_costs)
        self.japanese_bed_room_count.trace('w', self.update_additional_equipment_costs)
        self.western_room_count.trace('w', self.update_additional_equipment_costs)

        self.capacity_japanese_room.trace('w', self.update_additional_equipment_costs)
        self.capacity_japanese_western_room.trace('w', self.update_additional_equipment_costs)
        self.capacity_japanese_bed_room.trace('w', self.update_additional_equipment_costs)
        self.capacity_western_room.trace('w', self.update_additional_equipment_costs)

        # チェックボックスの変更を監視
        self.heat_source.trace('w', self.update_additional_equipment_subtotal)
        self.elevator.trace('w', self.update_additional_equipment_subtotal)
        self.gas_equipment.trace('w', self.update_additional_equipment_subtotal)
        self.septic_tank.trace('w', self.update_additional_equipment_subtotal)
        self.septic_tank.trace('w', self.update_additional_equipment_subtotal)
        self.central_ac.trace('w', self.update_additional_equipment_subtotal)
        self.fire_safety.trace('w', self.update_additional_equipment_subtotal)


        # コスト計算タブの客室数が変更されたら客室タブも更新
        self.japanese_room_count.trace('w', self.update_guest_arch_costs)
        self.japanese_western_room_count.trace('w', self.update_guest_arch_costs)
        self.japanese_bed_room_count.trace('w', self.update_guest_arch_costs)
        self.western_room_count.trace('w', self.update_guest_arch_costs)

        # スプリンクラー警告用のフラグ（同じ警告を複数回表示しないため）
        self.area_warning_shown = False # 面積警告用のフラグ（同じ警告を複数回表示しないため）
        self.area_warning_total_floor_area = 6000  # (5)イの　スプリンクラー　延べ床面積

        # 工事費率の金額表示用変数を初期化
        if not hasattr(self, 'direct_construction_cost_amount'):
            self.direct_construction_cost_amount = tk.StringVar(value="0")
        if not hasattr(self, 'direct_construction_cost_rate'):
            self.direct_construction_cost_rate = tk.StringVar(value="0.00")
        if not hasattr(self, 'net_construction_cost_amount'):
            self.net_construction_cost_amount = tk.StringVar(value="0")
        if not hasattr(self, 'site_management_cost_amount'):
            self.site_management_cost_amount = tk.StringVar(value="0")
        if not hasattr(self, 'common_temporary_cost_amount'):
            self.common_temporary_cost_amount = tk.StringVar(value="0")
        if not hasattr(self, 'general_admin_fee_amount'):
            self.general_admin_fee_amount = tk.StringVar(value="0")
        if not hasattr(self, 'lcc_region_rate_display'):
            self.lcc_region_rate_display = tk.StringVar(value="1.00")


        ###############################変数定義の範囲終了############################################
        # アプリケーション起動時、デフォルト値に基づいて初期単価計算を実行
        self.update_additional_equipment_costs()
        self.update_floor_area_inputs()

        # ★★★ 建設地域が変更されたときに単価計算を呼び出すバインド ★★★
        self.region.trace_add("write", self.update_construction_unit_price)
        # ★★★ 建物構造が変更されたときに単価計算を呼び出すバインド ★★★
        self.structure_type.trace_add("write", self.update_construction_unit_price)
        # ★★★ 仕様グレードが変更されたときに単価計算を呼び出すバインド ★★★
        self.selected_spec_grade.trace_add("write", self.update_construction_unit_price)

        self.update_construction_unit_price()

        # ========== ★ 物件管理機能用イベント設定 追加開始 ★ ==========
        # 物件リストの読み込み
        self.load_property_list()

        # ウィンドウを閉じる時の確認
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # タブ切り替え時のイベント（将来の拡張用）
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        # 主要な変数の変更検知（データ変更フラグ用）
        self.region.trace('w', self.mark_modified)
        self.structure_type.trace('w', self.mark_modified)
        self.floors.trace('w', self.mark_modified)
        # ========== ★ 物件管理機能用イベント設定 追加終了 ★ ==========


    def create_property_tab(self):
        """物件管理タブの作成"""
        canvas = tk.Canvas(self.property_frame)
        scrollbar = ttk.Scrollbar(self.property_frame, orient=tk.VERTICAL,
                                  command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>",
                              lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # タイトル
        ttk.Label(main_frame, text="物件管理",
                  font=title_font).grid(
            row=0, column=0, columnspan=3, pady=(0, 20), sticky=tk.W)

        row = 1

        # ★★★ ここから追加: 保存ボタンエリア ★★★
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))

        # ボタンを右寄せで配置

        ttk.Button(button_frame, text="💾 保存",
                   command=self.save_current_property).pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="📁 名前を付けて保存",
                   command=self.save_as_property).pack(side=tk.LEFT, padx=5)

        row += 1
        # ★★★ ここまで追加 ★★★
        # 物件一覧セクション
        list_frame = ttk.LabelFrame(main_frame, text="登録済み物件一覧", padding="15")
        list_frame.grid(row=row, column=0, columnspan=3,
                        sticky=(tk.W, tk.E), pady=(0, 20))

        # 物件リスト
        list_container = ttk.Frame(list_frame)
        list_container.pack(fill=tk.BOTH, expand=True)

        list_scroll = ttk.Scrollbar(list_container, orient=tk.VERTICAL)
        self.property_listbox = tk.Listbox(list_container, height=8,
                                           yscrollcommand=list_scroll.set,
                                           font=sheet_font)
        list_scroll.config(command=self.property_listbox.yview)

        self.property_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        list_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # リストボタン
        list_btn_frame = ttk.Frame(list_frame)
        list_btn_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(list_btn_frame, text="📂 開く",
                   command=self.load_selected_property).pack(side=tk.LEFT, padx=5)
        ttk.Button(list_btn_frame, text="🆕 新規作成",
                   command=self.create_new_property).pack(side=tk.LEFT, padx=5)
        ttk.Button(list_btn_frame, text="🗑️ 削除",
                   command=self.delete_property).pack(side=tk.LEFT, padx=5)

        row += 1

        # 物件情報入力セクション
        info_frame = ttk.LabelFrame(main_frame, text="物件基本情報", padding="15")
        info_frame.grid(row=row, column=0, columnspan=3,
                        sticky=(tk.W, tk.E), pady=(0, 20))

        # 物件名変数の初期化
        if not hasattr(self, 'property_name_var'):
            self.property_name_var = tk.StringVar(value="")
            self.property_address_var = tk.StringVar(value="")

        # 物件名
        ttk.Label(info_frame, text="物件名:",
                  font=sheet_font).grid(
            row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(info_frame, textvariable=self.property_name_var, width=50,
                  font=sheet_font).grid(row=0, column=1,
                                           sticky=(tk.W, tk.E), pady=5, padx=(10, 0))

        # 住所
        ttk.Label(info_frame, text="住所:", font=sheet_font).grid(
            row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(info_frame, textvariable=self.property_address_var,
                  width=50).grid(row=1, column=1,
                                 sticky=(tk.W, tk.E), pady=5, padx=(10, 0))

        # 備考
        ttk.Label(info_frame, text="備考:", font=sheet_font).grid(
            row=2, column=0, sticky=tk.NW, pady=5)
        notes_text = tk.Text(info_frame, width=50, height=5, font=sheet_font)
        notes_text.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        self.notes_text = notes_text

        info_frame.grid_columnconfigure(1, weight=1)

        row += 1

        # 使用方法の説明
        help_frame = ttk.LabelFrame(main_frame, text="使用方法", padding="15")
        help_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))

        help_text = """
1. 新規物件を作成: 「🆕 新規作成」ボタンをクリック
2. 物件情報を入力: 物件名、住所、備考を入力
3. 各タブでデータを入力: コストサマリ、温泉設備、レストランなど
4. 保存: 上部の「💾 保存」ボタンで保存
5. 既存物件を開く: 一覧から選択して「📂 開く」ボタン
        """
        ttk.Label(help_frame, text=help_text, font=sheet_font,
                  justify=tk.LEFT).pack(anchor=tk.W)

        # 変更検知
        self.property_name_var.trace('w', self.mark_modified)
        self.property_address_var.trace('w', self.mark_modified)

    def mark_modified(self, *args):
        """データ変更フラグを立てる"""
        self.data_modified = True


    def update_window_title(self, property_name=""):
        """ウィンドウタイトルを更新する

        Args:
            property_name (str): 表示する物件名（空文字列の場合はファイル名なしで表示）
        """
        if property_name:
            # 物件名がある場合: "システム名 - 物件名"
            self.root.title(f"{self.base_title} - {property_name}")
        else:
            # 物件名がない場合: システム名のみ
            self.root.title(self.base_title)


    def create_new_property(self):
        """新規物件の作成"""
        if self.data_modified:
            result = messagebox.askyesnocancel(
                "確認",
                "現在の物件に変更があります。保存しますか？")
            if result is None:  # キャンセル
                return
            elif result:  # はい
                self.save_current_property()

        # データをクリア
        self.current_property_id = None
        self.current_property_name.set("新規物件")
        self.property_name_var.set("")
        self.property_address_var.set("")
        self.notes_text.delete("1.0", tk.END)
        # ウィンドウタイトルをリセット
        self.update_window_title("")
        # 基本変数のリセット
        self.region.set("")
        self.base_date.set("")  # ★追加★
        self.completion_date.set("")  # ★追加★
        self.structure_type.set("")
        self.floors.set("")
        self.underground_floors.set("")

        # ★追加: 階ごとの面積データをクリア★
        if hasattr(self, 'floor_areas'):
            self.floor_areas.clear()
            self.update_floor_area_inputs()  # 入力欄を再生成

        # 客室数のリセット
        if hasattr(self, 'japanese_room_count'):
            self.japanese_room_count.set("0")
            self.japanese_western_room_count.set("0")
            self.japanese_bed_room_count.set("0")
            self.western_room_count.set("0")

        # レストラン建築
        if hasattr(self, 'restaurant_arch_restaurant_check'):
            self.restaurant_arch_restaurant_check.set(False)
            self.restaurant_arch_restaurant_area.set("0")
            self.restaurant_arch_livekitchen_check.set(False)
            self.restaurant_arch_livekitchen_area.set("0")
            self.restaurant_arch_kitchen_check.set(False)
            self.restaurant_arch_kitchen_area.set("0")

        # レストラン機械設備
        if hasattr(self, 'restaurant_mech_ac_check'):
            self.restaurant_mech_ac_check.set(False)
            self.restaurant_mech_ac_count.set("0")
            self.restaurant_mech_sprinkler_check.set(False)
            self.restaurant_mech_sprinkler_count.set("0")
            self.restaurant_mech_fire_hose_check.set(False)
            self.restaurant_mech_fire_hose_count.set("0")
            self.restaurant_mech_hood_check.set(False)
            self.restaurant_mech_hood_count.set("0")

        # レストラン電気設備
        if hasattr(self, 'restaurant_elec_led_check'):
            self.restaurant_elec_led_check.set(False)
            self.restaurant_elec_led_count.set("0")
            self.restaurant_elec_smoke_check.set(False)
            self.restaurant_elec_smoke_count.set("0")
            self.restaurant_elec_exit_light_check.set(False)
            self.restaurant_elec_exit_light_count.set("0")
            self.restaurant_elec_speaker_check.set(False)
            self.restaurant_elec_speaker_count.set("0")
            self.restaurant_elec_outlet_check.set(False)
            self.restaurant_elec_outlet_count.set("0")

        # レストラン厨房設備
        if hasattr(self, 'restaurant_kitchen_equipment_checks'):
            for check_var in self.restaurant_kitchen_equipment_checks:
                check_var.set(False)
            for count_var in self.restaurant_kitchen_equipment_counts:
                count_var.set("1")

        # レストラン家具
        if hasattr(self, 'restaurant_furniture_bar_counter_check'):
            self.restaurant_furniture_bar_counter_check.set(False)
            self.restaurant_furniture_bar_counter_count.set("0")
            self.restaurant_furniture_soft_drink_counter_check.set(False)
            self.restaurant_furniture_soft_drink_counter_count.set("0")
            self.restaurant_furniture_alcohol_counter_check.set(False)
            self.restaurant_furniture_alcohol_counter_count.set("0")
            self.restaurant_furniture_cutlery_counter_check.set(False)
            self.restaurant_furniture_cutlery_counter_count.set("0")
            self.restaurant_furniture_ice_counter_check.set(False)
            self.restaurant_furniture_ice_counter_count.set("0")
            self.restaurant_furniture_soft_cream_counter_check.set(False)
            self.restaurant_furniture_soft_cream_counter_count.set("0")
            self.restaurant_furniture_return_counter_check.set(False)
            self.restaurant_furniture_return_counter_count.set("0")

        # ラウンジ建築
        if hasattr(self, 'lounge_arch_front_check'):
            self.lounge_arch_front_check.set(False)
            self.lounge_arch_front_area.set("0")
            self.lounge_arch_lobby_check.set(False)
            self.lounge_arch_lobby_area.set("0")
            self.lounge_arch_facade_check.set(False)
            self.lounge_arch_facade_area.set("0")
            self.lounge_arch_shop_check.set(False)
            self.lounge_arch_shop_area.set("0")

        # ラウンジ機械設備
        if hasattr(self, 'lounge_mech_ac_check'):
            self.lounge_mech_ac_check.set(False)
            self.lounge_mech_ac_count.set("0")
            self.lounge_mech_sprinkler_check.set(False)
            self.lounge_mech_sprinkler_count.set("0")
            self.lounge_mech_fire_hose_check.set(False)
            self.lounge_mech_fire_hose_count.set("0")

        # ラウンジ電気設備
        if hasattr(self, 'lounge_elec_led_check'):
            self.lounge_elec_led_check.set(False)
            self.lounge_elec_led_count.set("0")
            self.lounge_elec_smoke_check.set(False)
            self.lounge_elec_smoke_count.set("0")
            self.lounge_elec_exit_light_check.set(False)
            self.lounge_elec_exit_light_count.set("0")
            self.lounge_elec_speaker_check.set(False)
            self.lounge_elec_speaker_count.set("0")
            self.lounge_elec_outlet_check.set(False)
            self.lounge_elec_outlet_count.set("0")

        # アミューズメント建築
        if hasattr(self, 'amusement_arch_pingpong_check'):
            self.amusement_arch_pingpong_check.set(False)
            self.amusement_arch_pingpong_area.set("0")
            self.amusement_arch_kids_check.set(False)
            self.amusement_arch_kids_area.set("0")
            self.amusement_arch_manga_check.set(False)
            self.amusement_arch_manga_area.set("0")

        # アミューズメント機械設備
        if hasattr(self, 'amusement_mech_ac_check'):
            self.amusement_mech_ac_check.set(False)
            self.amusement_mech_ac_count.set("0")
            self.amusement_mech_sprinkler_check.set(False)
            self.amusement_mech_sprinkler_count.set("0")
            self.amusement_mech_fire_hose_check.set(False)
            self.amusement_mech_fire_hose_count.set("0")

        # アミューズメント電気設備
        if hasattr(self, 'amusement_elec_led_check'):
            self.amusement_elec_led_check.set(False)
            self.amusement_elec_led_count.set("0")
            self.amusement_elec_smoke_check.set(False)
            self.amusement_elec_smoke_count.set("0")
            self.amusement_elec_exit_light_check.set(False)
            self.amusement_elec_exit_light_count.set("0")
            self.amusement_elec_speaker_check.set(False)
            self.amusement_elec_speaker_count.set("0")
            self.amusement_elec_outlet_check.set(False)
            self.amusement_elec_outlet_count.set("0")

        # 客室定員
        if hasattr(self, 'capacity_japanese_room'):
            self.capacity_japanese_room.set("4.5")
            self.capacity_japanese_western_room.set("4.0")
            self.capacity_japanese_bed_room.set("3.0")
            self.capacity_western_room.set("3.0")

        # 客室建築
        if hasattr(self, 'guest_arch_japanese_check'):
            self.guest_arch_japanese_check.set(False)
            self.guest_arch_japanese_area.set("0")
            self.guest_arch_japanese_western_check.set(False)
            self.guest_arch_japanese_western_area.set("0")
            self.guest_arch_japanese_bed_check.set(False)
            self.guest_arch_japanese_bed_area.set("0")
            self.guest_arch_western_check.set(False)
            self.guest_arch_western_area.set("0")

        # 客室機械設備
        if hasattr(self, 'guest_mech_ac_check'):
            self.guest_mech_ac_check.set(False)
            self.guest_mech_ac_count.set("0")
            self.guest_mech_wash_basin_check.set(False)
            self.guest_mech_wash_basin_count.set("0")
            self.guest_mech_sprinkler_check.set(False)
            self.guest_mech_sprinkler_count.set("0")

        # 客室電気設備
        if hasattr(self, 'guest_elec_main_light_check'):
            self.guest_elec_main_light_check.set(False)
            self.guest_elec_main_light_count.set("0")
            self.guest_elec_smoke_check.set(False)
            self.guest_elec_smoke_count.set("0")
            self.guest_elec_heat_detector_check.set(False)
            self.guest_elec_heat_detector_count.set("0")
            self.guest_elec_speaker_check.set(False)
            self.guest_elec_speaker_count.set("0")
            self.guest_elec_outlet_check.set(False)
            self.guest_elec_outlet_count.set("0")

        # 客室設備
        if hasattr(self, 'guest_room_tv_cabinet'):
            self.guest_room_tv_cabinet.set(False)
            self.guest_room_headboard.set(False)
            self.guest_private_open_bath.set(False)


        self.data_modified = False
        messagebox.showinfo("完了", "新規物件を作成しました")

    def save_current_property(self):
        """現在の物件を保存"""
        property_name = self.property_name_var.get().strip()

        if not property_name:
            messagebox.showerror("エラー", "物件名を入力してください")
            return

        # 物件IDの生成（新規の場合）
        if self.current_property_id is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.current_property_id = f"prop_{timestamp}"

        # データの収集
        data = self.collect_all_data()

        # ファイルに保存
        filename = os.path.join(self.data_dir, f"{self.current_property_id}.json")
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            self.current_property_name.set(property_name)
            self.update_window_title(property_name)
            self.data_modified = False
            self.load_property_list()
            messagebox.showinfo("完了", f"物件「{property_name}」を保存しました")
        except Exception as e:
            messagebox.showerror("エラー", f"保存に失敗しました: {str(e)}")

    def save_as_property(self):
        """名前を付けて保存"""
        self.current_property_id = None
        self.save_current_property()

    def collect_all_data(self):
        """全データを収集"""
        data = {
            'property_id': self.current_property_id,
            'property_name': self.property_name_var.get(),
            'property_address': self.property_address_var.get(),
            'property_notes': self.notes_text.get("1.0", tk.END).strip(),
            'saved_at': datetime.now().isoformat(),
        }

        # 基本情報
        data['basic_info'] = {
            'region': self.region.get(),
            'base_date': self.base_date.get(),  # ★追加★
            'completion_date': self.completion_date.get(),  # ★修正（if文削除）★
            'structure_type': self.structure_type.get(),
            'floors': self.floors.get(),
            'underground_floors': self.underground_floors.get(),
            'spec_grade': self.spec_grade.get() if hasattr(self, 'spec_grade') else "",
        }

        # ★追加: 階ごとの面積データ★
        data['floor_areas'] = {}
        if hasattr(self, 'floor_areas'):
            for floor_key, floor_var in self.floor_areas.items():
                try:
                    area_value = floor_var.get()
                    data['floor_areas'][floor_key] = area_value
                except:
                    data['floor_areas'][floor_key] = "0"

        # 客室数
        if hasattr(self, 'japanese_room_count'):
            data['room_counts'] = {
                'japanese': self.japanese_room_count.get(),
                'japanese_western': self.japanese_western_room_count.get(),
                'japanese_bed': self.japanese_bed_room_count.get(),
                'western': self.western_room_count.get(),
            }

        # 温泉設備データ
        data['onsen_data'] = {
            'settings': {
                'bath_open_hours': self.bath_open_hours.get(),
                'wash_time': self.wash_time.get(),
                'tub_time': self.tub_time.get(),
                'peak_factor': self.peak_factor.get(),
                'men_ratio': self.men_ratio.get(),
                'women_ratio': self.women_ratio.get(),
            },
            'men': {
                'indoor_bath_check': self.men_indoor_bath_check.get(),
                'indoor_bath_area': self.men_indoor_bath_area.get(),
                'outdoor_bath_check': self.men_outdoor_bath_check.get(),
                'outdoor_bath_area': self.men_outdoor_bath_area.get(),
                'bathroom_check': self.men_bathroom_check.get(),
                'bathroom_area': self.men_bathroom_area.get(),
                'partition_count': self.men_partition_count.get(),
                'wash_lighting_count': self.men_wash_lighting_count.get(),
                'shower_faucet_count': self.men_shower_faucet_count.get(),
                'sauna': self.men_sauna.get(),
                'jacuzzi': self.men_jacuzzi.get(),
                'water_bath': self.men_water_bath.get(),
            },
            'women': {
                'indoor_bath_check': self.women_indoor_bath_check.get(),
                'indoor_bath_area': self.women_indoor_bath_area.get(),
                'outdoor_bath_check': self.women_outdoor_bath_check.get(),
                'outdoor_bath_area': self.women_outdoor_bath_area.get(),
                'bathroom_check': self.women_bathroom_check.get(),
                'bathroom_area': self.women_bathroom_area.get(),
                'partition_count': self.women_partition_count.get(),
                'wash_lighting_count': self.women_wash_lighting_count.get(),
                'shower_faucet_count': self.women_shower_faucet_count.get(),
                'sauna': self.women_sauna.get(),
                'jacuzzi': self.women_jacuzzi.get(),
                'water_bath': self.women_water_bath.get(),
            }
        }

        # レストランデータ
        data['restaurant_data'] = {
            'arch': {
                'restaurant_check': self.restaurant_arch_restaurant_check.get(),
                'restaurant_area': self.restaurant_arch_restaurant_area.get(),
                'livekitchen_check': self.restaurant_arch_livekitchen_check.get(),
                'livekitchen_area': self.restaurant_arch_livekitchen_area.get(),
                'kitchen_check': self.restaurant_arch_kitchen_check.get(),
                'kitchen_area': self.restaurant_arch_kitchen_area.get(),
            },
            'mech': {
                'ac_check': self.restaurant_mech_ac_check.get(),
                'ac_count': self.restaurant_mech_ac_count.get(),
                'sprinkler_check': self.restaurant_mech_sprinkler_check.get(),
                'sprinkler_count': self.restaurant_mech_sprinkler_count.get(),
                'fire_hose_check': self.restaurant_mech_fire_hose_check.get(),
                'fire_hose_count': self.restaurant_mech_fire_hose_count.get(),
                'hood_check': self.restaurant_mech_hood_check.get(),
                'hood_count': self.restaurant_mech_hood_count.get(),
            },
            'elec': {
                'led_check': self.restaurant_elec_led_check.get(),
                'led_count': self.restaurant_elec_led_count.get(),
                'smoke_check': self.restaurant_elec_smoke_check.get(),
                'smoke_count': self.restaurant_elec_smoke_count.get(),
                'exit_light_check': self.restaurant_elec_exit_light_check.get(),
                'exit_light_count': self.restaurant_elec_exit_light_count.get(),
                'speaker_check': self.restaurant_elec_speaker_check.get(),
                'speaker_count': self.restaurant_elec_speaker_count.get(),
                'outlet_check': self.restaurant_elec_outlet_check.get(),
                'outlet_count': self.restaurant_elec_outlet_count.get(),
            },
            'kitchen_equipment': [
                {
                    'check': check_var.get(),
                    'count': count_var.get()
                }
                for check_var, count_var in zip(
                    self.restaurant_kitchen_equipment_checks,
                    self.restaurant_kitchen_equipment_counts
                )
            ],
            'furniture': {
                'bar_counter_check': self.restaurant_furniture_bar_counter_check.get(),
                'bar_counter_count': self.restaurant_furniture_bar_counter_count.get(),
                'soft_drink_counter_check': self.restaurant_furniture_soft_drink_counter_check.get(),
                'soft_drink_counter_count': self.restaurant_furniture_soft_drink_counter_count.get(),
                'alcohol_counter_check': self.restaurant_furniture_alcohol_counter_check.get(),
                'alcohol_counter_count': self.restaurant_furniture_alcohol_counter_count.get(),
                'cutlery_counter_check': self.restaurant_furniture_cutlery_counter_check.get(),
                'cutlery_counter_count': self.restaurant_furniture_cutlery_counter_count.get(),
                'ice_counter_check': self.restaurant_furniture_ice_counter_check.get(),
                'ice_counter_count': self.restaurant_furniture_ice_counter_count.get(),
                'soft_cream_counter_check': self.restaurant_furniture_soft_cream_counter_check.get(),
                'soft_cream_counter_count': self.restaurant_furniture_soft_cream_counter_count.get(),
                'return_counter_check': self.restaurant_furniture_return_counter_check.get(),
                'return_counter_count': self.restaurant_furniture_return_counter_count.get(),
            }
        }

        # ラウンジデータ
        data['lounge_data'] = {
            'arch': {
                'front_check': self.lounge_arch_front_check.get(),
                'front_area': self.lounge_arch_front_area.get(),
                'lobby_check': self.lounge_arch_lobby_check.get(),
                'lobby_area': self.lounge_arch_lobby_area.get(),
                'facade_check': self.lounge_arch_facade_check.get(),
                'facade_area': self.lounge_arch_facade_area.get(),
                'shop_check': self.lounge_arch_shop_check.get(),
                'shop_area': self.lounge_arch_shop_area.get(),
            },
            'mech': {
                'ac_check': self.lounge_mech_ac_check.get(),
                'ac_count': self.lounge_mech_ac_count.get(),
                'sprinkler_check': self.lounge_mech_sprinkler_check.get(),
                'sprinkler_count': self.lounge_mech_sprinkler_count.get(),
                'fire_hose_check': self.lounge_mech_fire_hose_check.get(),
                'fire_hose_count': self.lounge_mech_fire_hose_count.get(),
            },
            'elec': {
                'led_check': self.lounge_elec_led_check.get(),
                'led_count': self.lounge_elec_led_count.get(),
                'smoke_check': self.lounge_elec_smoke_check.get(),
                'smoke_count': self.lounge_elec_smoke_count.get(),
                'exit_light_check': self.lounge_elec_exit_light_check.get(),
                'exit_light_count': self.lounge_elec_exit_light_count.get(),
                'speaker_check': self.lounge_elec_speaker_check.get(),
                'speaker_count': self.lounge_elec_speaker_count.get(),
                'outlet_check': self.lounge_elec_outlet_check.get(),
                'outlet_count': self.lounge_elec_outlet_count.get(),
            }
        }

        # アミューズメントデータ
        data['amusement_data'] = {
            'arch': {
                'pingpong_check': self.amusement_arch_pingpong_check.get(),
                'pingpong_area': self.amusement_arch_pingpong_area.get(),
                'kids_check': self.amusement_arch_kids_check.get(),
                'kids_area': self.amusement_arch_kids_area.get(),
                'manga_check': self.amusement_arch_manga_check.get(),
                'manga_area': self.amusement_arch_manga_area.get(),
            },
            # ★★★ ここから追加 ★★★
            'mech': {
                'ac_check': self.amusement_mech_ac_check.get(),
                'ac_count': self.amusement_mech_ac_count.get(),
                'sprinkler_check': self.amusement_mech_sprinkler_check.get(),
                'sprinkler_count': self.amusement_mech_sprinkler_count.get(),
                'fire_hose_check': self.amusement_mech_fire_hose_check.get(),
                'fire_hose_count': self.amusement_mech_fire_hose_count.get(),
            },
            'elec': {
                'led_check': self.amusement_elec_led_check.get(),
                'led_count': self.amusement_elec_led_count.get(),
                'smoke_check': self.amusement_elec_smoke_check.get(),
                'smoke_count': self.amusement_elec_smoke_count.get(),
                'exit_light_check': self.amusement_elec_exit_light_check.get(),
                'exit_light_count': self.amusement_elec_exit_light_count.get(),
                'speaker_check': self.amusement_elec_speaker_check.get(),
                'speaker_count': self.amusement_elec_speaker_count.get(),
                'outlet_check': self.amusement_elec_outlet_check.get(),
                'outlet_count': self.amusement_elec_outlet_count.get(),
            }
            # ★★★ ここまで追加 ★★★
        }

        # 通路データ
        data['hallway_data'] = {
            'arch': {},
            'mech': {},
            'elec': {}
        }

        # 通路建築項目（各階の面積）
        if hasattr(self, 'hallway_arch_checks'):
            for floor_key in self.hallway_arch_checks.keys():
                data['hallway_data']['arch'][floor_key] = {
                    'check': self.hallway_arch_checks[floor_key].get(),
                    'area': self.hallway_arch_areas[floor_key].get(),
                    'cost': self.hallway_arch_costs[floor_key].get()
                }

        # 通路機械設備
        if hasattr(self, 'hallway_mech_ac_check'):
            data['hallway_data']['mech'] = {
                'ac': {
                    'check': self.hallway_mech_ac_check.get(),
                    'count': self.hallway_mech_ac_count.get(),
                    'cost': self.hallway_mech_ac_cost.get()
                },
                'sprinkler': {
                    'check': self.hallway_mech_sprinkler_check.get(),
                    'count': self.hallway_mech_sprinkler_count.get(),
                    'cost': self.hallway_mech_sprinkler_cost.get()
                },
                'fire_hose': {
                    'check': self.hallway_mech_fire_hose_check.get(),
                    'count': self.hallway_mech_fire_hose_count.get(),
                    'cost': self.hallway_mech_fire_hose_cost.get()
                }
            }

        # 通路電気設備
        if hasattr(self, 'hallway_elec_led_check'):
            data['hallway_data']['elec'] = {
                'led': {
                    'check': self.hallway_elec_led_check.get(),
                    'count': self.hallway_elec_led_count.get(),
                    'cost': self.hallway_elec_led_cost.get()
                },
                'smoke': {
                    'check': self.hallway_elec_smoke_check.get(),
                    'count': self.hallway_elec_smoke_count.get(),
                    'cost': self.hallway_elec_smoke_cost.get()
                },
                'exit_light': {
                    'check': self.hallway_elec_exit_light_check.get(),
                    'count': self.hallway_elec_exit_light_count.get(),
                    'cost': self.hallway_elec_exit_light_cost.get()
                },
                'speaker': {
                    'check': self.hallway_elec_speaker_check.get(),
                    'count': self.hallway_elec_speaker_count.get(),
                    'cost': self.hallway_elec_speaker_cost.get()
                },
                'outlet': {
                    'check': self.hallway_elec_outlet_check.get(),
                    'count': self.hallway_elec_outlet_count.get(),
                    'cost': self.hallway_elec_outlet_cost.get()
                }
            }

        # 客室データ
        data['guest_room_data'] = {
            'capacity': {
                'japanese': self.capacity_japanese_room.get(),
                'japanese_western': self.capacity_japanese_western_room.get(),
                'japanese_bed': self.capacity_japanese_bed_room.get(),
                'western': self.capacity_western_room.get(),
            },
            'arch': {
                'japanese_check': self.guest_arch_japanese_check.get(),
                'japanese_area': self.guest_arch_japanese_area.get(),
                'japanese_western_check': self.guest_arch_japanese_western_check.get(),
                'japanese_western_area': self.guest_arch_japanese_western_area.get(),
                'japanese_bed_check': self.guest_arch_japanese_bed_check.get(),
                'japanese_bed_area': self.guest_arch_japanese_bed_area.get(),
                'western_check': self.guest_arch_western_check.get(),
                'western_area': self.guest_arch_western_area.get(),
            },
            # ★★★ ここから追加 ★★★
            'mech': {
                'ac_check': self.guest_mech_ac_check.get(),
                'ac_count': self.guest_mech_ac_count.get(),
                'wash_basin_check': self.guest_mech_wash_basin_check.get(),
                'wash_basin_count': self.guest_mech_wash_basin_count.get(),
                'sprinkler_check': self.guest_mech_sprinkler_check.get(),
                'sprinkler_count': self.guest_mech_sprinkler_count.get(),
            },
            'elec': {
                'main_light_check': self.guest_elec_main_light_check.get(),
                'main_light_count': self.guest_elec_main_light_count.get(),
                'smoke_check': self.guest_elec_smoke_check.get(),
                'smoke_count': self.guest_elec_smoke_count.get(),
                'heat_detector_check': self.guest_elec_heat_detector_check.get(),
                'heat_detector_count': self.guest_elec_heat_detector_count.get(),
                'speaker_check': self.guest_elec_speaker_check.get(),
                'speaker_count': self.guest_elec_speaker_count.get(),
                'outlet_check': self.guest_elec_outlet_check.get(),
                'outlet_count': self.guest_elec_outlet_count.get(),
            },
            'equipment': {
                'tv_cabinet': self.guest_room_tv_cabinet.get(),
                'headboard': self.guest_room_headboard.get(),
                'private_open_bath': self.guest_private_open_bath.get(),
            }
            # ★★★ ここまで追加 ★★★
        }

        # 付加設備
        data['additional_equipment'] = {
            'heat_source': self.heat_source.get(),
            'elevator': self.elevator.get(),
            'gas_equipment': self.gas_equipment.get(),
            'septic_tank': self.septic_tank.get(),
            'central_ac': self.central_ac.get(),
            'fire_safety': self.fire_safety.get(),
            'generator': self.generator.get() if hasattr(self, 'generator') else False,
        }

        return data

    def load_property_list(self):
        """物件リストの読み込み"""
        self.property_listbox.delete(0, tk.END)

        if not os.path.exists(self.data_dir):
            return

        files = [f for f in os.listdir(self.data_dir) if f.endswith('.json')]

        for file in sorted(files, reverse=True):
            filepath = os.path.join(self.data_dir, file)
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    property_name = data.get('property_name', '名称未設定')
                    saved_at = data.get('saved_at', '')
                    if saved_at:
                        saved_date = saved_at[:10]
                        display_text = f"{property_name} ({saved_date})"
                    else:
                        display_text = property_name

                    self.property_listbox.insert(tk.END, display_text)
                    self.properties_data[display_text] = (file, data)
            except:
                continue

    def load_selected_property(self):
        """選択された物件を読み込み"""
        selection = self.property_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "物件を選択してください")
            return

        if self.data_modified:
            result = messagebox.askyesnocancel(
                "確認",
                "現在の物件に変更があります。保存しますか？")
            if result is None:
                return
            elif result:
                self.save_current_property()

        selected_text = self.property_listbox.get(selection[0])
        file, data = self.properties_data[selected_text]

        self.load_property_data(data)
        self.data_modified = False
        messagebox.showinfo("完了", f"物件「{data['property_name']}」を読み込みました")

    def load_property_data(self, data):
        """物件データをフォームに読み込み"""
        self.current_property_id = data.get('property_id')
        self.current_property_name.set(data.get('property_name', ''))
        self.property_name_var.set(data.get('property_name', ''))
        self.property_address_var.set(data.get('property_address', ''))
        self.update_window_title(data.get('property_name', ''))
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", data.get('property_notes', ''))

        # 基本情報
        basic_info = data.get('basic_info', {})
        self.region.set(basic_info.get('region', ''))
        self.base_date.set(basic_info.get('base_date', ''))  # ★追加★
        self.completion_date.set(basic_info.get('completion_date', ''))  # ★追加（if文なし）★
        self.structure_type.set(basic_info.get('structure_type', ''))
        self.floors.set(basic_info.get('floors', ''))
        self.underground_floors.set(basic_info.get('underground_floors', ''))
        if hasattr(self, 'spec_grade'):
            self.spec_grade.set(basic_info.get('spec_grade', ''))

        # ★追加: 階ごとの面積データの復元★
        floor_areas_data = data.get('floor_areas', {})
        if hasattr(self, 'floor_areas'):
            # まず既存の階数入力に基づいて入力欄を生成
            self.update_floor_area_inputs()

            # 次に保存されていた値を設定
            for floor_key, area_value in floor_areas_data.items():
                if floor_key in self.floor_areas:
                    self.floor_areas[floor_key].set(area_value)

        # 合計面積を更新
        if hasattr(self, 'update_total_floor_areas'):
            self.update_total_floor_areas()

        # 客室数
        room_counts = data.get('room_counts', {})
        if hasattr(self, 'japanese_room_count'):
            self.japanese_room_count.set(room_counts.get('japanese', '0'))
            self.japanese_western_room_count.set(room_counts.get('japanese_western', '0'))
            self.japanese_bed_room_count.set(room_counts.get('japanese_bed', '0'))
            self.western_room_count.set(room_counts.get('western', '0'))

        # 温泉設備データ
        onsen_data = data.get('onsen_data', {})
        settings = onsen_data.get('settings', {})
        self.bath_open_hours.set(settings.get('bath_open_hours', 6.0))
        self.wash_time.set(settings.get('wash_time', 10.0))
        self.tub_time.set(settings.get('tub_time', 20.0))
        self.peak_factor.set(settings.get('peak_factor', 4.0))
        self.men_ratio.set(settings.get('men_ratio', 55.0))
        self.women_ratio.set(settings.get('women_ratio', 45.0))

        men_data = onsen_data.get('men', {})
        self.men_indoor_bath_check.set(men_data.get('indoor_bath_check', False))
        self.men_indoor_bath_area.set(men_data.get('indoor_bath_area', '0'))
        self.men_outdoor_bath_check.set(men_data.get('outdoor_bath_check', False))
        self.men_outdoor_bath_area.set(men_data.get('outdoor_bath_area', '0'))
        self.men_bathroom_check.set(men_data.get('bathroom_check', False))
        self.men_bathroom_area.set(men_data.get('bathroom_area', '0'))
        self.men_partition_count.set(men_data.get('partition_count', '0'))
        self.men_wash_lighting_count.set(men_data.get('wash_lighting_count', '0'))
        self.men_shower_faucet_count.set(men_data.get('shower_faucet_count', '0'))
        self.men_sauna.set(men_data.get('sauna', False))
        self.men_jacuzzi.set(men_data.get('jacuzzi', False))
        self.men_water_bath.set(men_data.get('water_bath', False))

        women_data = onsen_data.get('women', {})
        self.women_indoor_bath_check.set(women_data.get('indoor_bath_check', False))
        self.women_indoor_bath_area.set(women_data.get('indoor_bath_area', '0'))
        self.women_outdoor_bath_check.set(women_data.get('outdoor_bath_check', False))
        self.women_outdoor_bath_area.set(women_data.get('outdoor_bath_area', '0'))
        self.women_bathroom_check.set(women_data.get('bathroom_check', False))
        self.women_bathroom_area.set(women_data.get('bathroom_area', '0'))
        self.women_partition_count.set(women_data.get('partition_count', '0'))
        self.women_wash_lighting_count.set(women_data.get('wash_lighting_count', '0'))
        self.women_shower_faucet_count.set(women_data.get('shower_faucet_count', '0'))
        self.women_sauna.set(women_data.get('sauna', False))
        self.women_jacuzzi.set(women_data.get('jacuzzi', False))
        self.women_water_bath.set(women_data.get('water_bath', False))

        # レストランデータ
        restaurant_data = data.get('restaurant_data', {})
        arch = restaurant_data.get('arch', {})
        self.restaurant_arch_restaurant_check.set(arch.get('restaurant_check', False))
        self.restaurant_arch_restaurant_area.set(arch.get('restaurant_area', '0'))
        self.restaurant_arch_livekitchen_check.set(arch.get('livekitchen_check', False))
        self.restaurant_arch_livekitchen_area.set(arch.get('livekitchen_area', '0'))
        self.restaurant_arch_kitchen_check.set(arch.get('kitchen_check', False))
        self.restaurant_arch_kitchen_area.set(arch.get('kitchen_area', '0'))

        # レストラン機械設備
        mech = restaurant_data.get('mech', {})
        self.restaurant_mech_ac_check.set(mech.get('ac_check', False))
        self.restaurant_mech_ac_count.set(mech.get('ac_count', '0'))
        self.restaurant_mech_sprinkler_check.set(mech.get('sprinkler_check', False))
        self.restaurant_mech_sprinkler_count.set(mech.get('sprinkler_count', '0'))
        self.restaurant_mech_fire_hose_check.set(mech.get('fire_hose_check', False))
        self.restaurant_mech_fire_hose_count.set(mech.get('fire_hose_count', '0'))
        self.restaurant_mech_hood_check.set(mech.get('hood_check', False))
        self.restaurant_mech_hood_count.set(mech.get('hood_count', '0'))

        # レストラン電気設備
        elec = restaurant_data.get('elec', {})
        self.restaurant_elec_led_check.set(elec.get('led_check', False))
        self.restaurant_elec_led_count.set(elec.get('led_count', '0'))
        self.restaurant_elec_smoke_check.set(elec.get('smoke_check', False))
        self.restaurant_elec_smoke_count.set(elec.get('smoke_count', '0'))
        self.restaurant_elec_exit_light_check.set(elec.get('exit_light_check', False))
        self.restaurant_elec_exit_light_count.set(elec.get('exit_light_count', '0'))
        self.restaurant_elec_speaker_check.set(elec.get('speaker_check', False))
        self.restaurant_elec_speaker_count.set(elec.get('speaker_count', '0'))
        self.restaurant_elec_outlet_check.set(elec.get('outlet_check', False))
        self.restaurant_elec_outlet_count.set(elec.get('outlet_count', '0'))

        # レストラン厨房設備
        kitchen_equipment = restaurant_data.get('kitchen_equipment', [])
        for idx, item in enumerate(kitchen_equipment):
            if idx < len(self.restaurant_kitchen_equipment_checks):
                self.restaurant_kitchen_equipment_checks[idx].set(item.get('check', False))
                self.restaurant_kitchen_equipment_counts[idx].set(item.get('count', '1'))

        # レストラン家具
        furniture = restaurant_data.get('furniture', {})
        self.restaurant_furniture_bar_counter_check.set(furniture.get('bar_counter_check', False))
        self.restaurant_furniture_bar_counter_count.set(furniture.get('bar_counter_count', '0'))
        self.restaurant_furniture_soft_drink_counter_check.set(furniture.get('soft_drink_counter_check', False))
        self.restaurant_furniture_soft_drink_counter_count.set(furniture.get('soft_drink_counter_count', '0'))
        self.restaurant_furniture_alcohol_counter_check.set(furniture.get('alcohol_counter_check', False))
        self.restaurant_furniture_alcohol_counter_count.set(furniture.get('alcohol_counter_count', '0'))
        self.restaurant_furniture_cutlery_counter_check.set(furniture.get('cutlery_counter_check', False))
        self.restaurant_furniture_cutlery_counter_count.set(furniture.get('cutlery_counter_count', '0'))
        self.restaurant_furniture_ice_counter_check.set(furniture.get('ice_counter_check', False))
        self.restaurant_furniture_ice_counter_count.set(furniture.get('ice_counter_count', '0'))
        self.restaurant_furniture_soft_cream_counter_check.set(furniture.get('soft_cream_counter_check', False))
        self.restaurant_furniture_soft_cream_counter_count.set(furniture.get('soft_cream_counter_count', '0'))
        self.restaurant_furniture_return_counter_check.set(furniture.get('return_counter_check', False))
        self.restaurant_furniture_return_counter_count.set(furniture.get('return_counter_count', '0'))

        # ラウンジデータ
        lounge_data = data.get('lounge_data', {})
        arch = lounge_data.get('arch', {})
        self.lounge_arch_front_check.set(arch.get('front_check', False))
        self.lounge_arch_front_area.set(arch.get('front_area', '0'))
        self.lounge_arch_lobby_check.set(arch.get('lobby_check', False))
        self.lounge_arch_lobby_area.set(arch.get('lobby_area', '0'))
        self.lounge_arch_facade_check.set(arch.get('facade_check', False))
        self.lounge_arch_facade_area.set(arch.get('facade_area', '0'))
        self.lounge_arch_shop_check.set(arch.get('shop_check', False))
        self.lounge_arch_shop_area.set(arch.get('shop_area', '0'))
        # ラウンジ機械設備
        mech = lounge_data.get('mech', {})
        self.lounge_mech_ac_check.set(mech.get('ac_check', False))
        self.lounge_mech_ac_count.set(mech.get('ac_count', '0'))
        self.lounge_mech_sprinkler_check.set(mech.get('sprinkler_check', False))
        self.lounge_mech_sprinkler_count.set(mech.get('sprinkler_count', '0'))
        self.lounge_mech_fire_hose_check.set(mech.get('fire_hose_check', False))
        self.lounge_mech_fire_hose_count.set(mech.get('fire_hose_count', '0'))
        # ラウンジ電気設備
        elec = lounge_data.get('elec', {})
        self.lounge_elec_led_check.set(elec.get('led_check', False))
        self.lounge_elec_led_count.set(elec.get('led_count', '0'))
        self.lounge_elec_smoke_check.set(elec.get('smoke_check', False))
        self.lounge_elec_smoke_count.set(elec.get('smoke_count', '0'))
        self.lounge_elec_exit_light_check.set(elec.get('exit_light_check', False))
        self.lounge_elec_exit_light_count.set(elec.get('exit_light_count', '0'))
        self.lounge_elec_speaker_check.set(elec.get('speaker_check', False))
        self.lounge_elec_speaker_count.set(elec.get('speaker_count', '0'))
        self.lounge_elec_outlet_check.set(elec.get('outlet_check', False))
        self.lounge_elec_outlet_count.set(elec.get('outlet_count', '0'))
        # アミューズメントデータ
        amusement_data = data.get('amusement_data', {})
        arch = amusement_data.get('arch', {})
        self.amusement_arch_pingpong_check.set(arch.get('pingpong_check', False))
        self.amusement_arch_pingpong_area.set(arch.get('pingpong_area', '0'))
        self.amusement_arch_kids_check.set(arch.get('kids_check', False))
        self.amusement_arch_kids_area.set(arch.get('kids_area', '0'))
        self.amusement_arch_manga_check.set(arch.get('manga_check', False))
        self.amusement_arch_manga_area.set(arch.get('manga_area', '0'))

        # アミューズメント機械設備
        mech = amusement_data.get('mech', {})
        self.amusement_mech_ac_check.set(mech.get('ac_check', False))
        self.amusement_mech_ac_count.set(mech.get('ac_count', '0'))
        self.amusement_mech_sprinkler_check.set(mech.get('sprinkler_check', False))
        self.amusement_mech_sprinkler_count.set(mech.get('sprinkler_count', '0'))
        self.amusement_mech_fire_hose_check.set(mech.get('fire_hose_check', False))
        self.amusement_mech_fire_hose_count.set(mech.get('fire_hose_count', '0'))

        # アミューズメント電気設備
        elec = amusement_data.get('elec', {})
        self.amusement_elec_led_check.set(elec.get('led_check', False))
        self.amusement_elec_led_count.set(elec.get('led_count', '0'))
        self.amusement_elec_smoke_check.set(elec.get('smoke_check', False))
        self.amusement_elec_smoke_count.set(elec.get('smoke_count', '0'))
        self.amusement_elec_exit_light_check.set(elec.get('exit_light_check', False))
        self.amusement_elec_exit_light_count.set(elec.get('exit_light_count', '0'))
        self.amusement_elec_speaker_check.set(elec.get('speaker_check', False))
        self.amusement_elec_speaker_count.set(elec.get('speaker_count', '0'))
        self.amusement_elec_outlet_check.set(elec.get('outlet_check', False))
        self.amusement_elec_outlet_count.set(elec.get('outlet_count', '0'))

        # 通路データ
        hallway_data = data.get('hallway_data', {})

        # 通路建築項目（各階の面積）
        if hasattr(self, 'hallway_arch_checks'):
            arch = hallway_data.get('arch', {})
            for floor_key, floor_data in arch.items():
                if floor_key in self.hallway_arch_checks:
                    self.hallway_arch_checks[floor_key].set(floor_data.get('check', False))
                    self.hallway_arch_areas[floor_key].set(floor_data.get('area', '0'))
                    self.hallway_arch_costs[floor_key].set(floor_data.get('cost', '0'))

        # 通路機械設備
        if hasattr(self, 'hallway_mech_ac_check'):
            mech = hallway_data.get('mech', {})

            ac = mech.get('ac', {})
            self.hallway_mech_ac_check.set(ac.get('check', False))
            self.hallway_mech_ac_count.set(ac.get('count', '0'))
            self.hallway_mech_ac_cost.set(ac.get('cost', '0'))

            sprinkler = mech.get('sprinkler', {})
            self.hallway_mech_sprinkler_check.set(sprinkler.get('check', False))
            self.hallway_mech_sprinkler_count.set(sprinkler.get('count', '0'))
            self.hallway_mech_sprinkler_cost.set(sprinkler.get('cost', '0'))

            fire_hose = mech.get('fire_hose', {})
            self.hallway_mech_fire_hose_check.set(fire_hose.get('check', False))
            self.hallway_mech_fire_hose_count.set(fire_hose.get('count', '0'))
            self.hallway_mech_fire_hose_cost.set(fire_hose.get('cost', '0'))

        # 通路電気設備
        if hasattr(self, 'hallway_elec_led_check'):
            elec = hallway_data.get('elec', {})

            led = elec.get('led', {})
            self.hallway_elec_led_check.set(led.get('check', False))
            self.hallway_elec_led_count.set(led.get('count', '0'))
            self.hallway_elec_led_cost.set(led.get('cost', '0'))

            smoke = elec.get('smoke', {})
            self.hallway_elec_smoke_check.set(smoke.get('check', False))
            self.hallway_elec_smoke_count.set(smoke.get('count', '0'))
            self.hallway_elec_smoke_cost.set(smoke.get('cost', '0'))

            exit_light = elec.get('exit_light', {})
            self.hallway_elec_exit_light_check.set(exit_light.get('check', False))
            self.hallway_elec_exit_light_count.set(exit_light.get('count', '0'))
            self.hallway_elec_exit_light_cost.set(exit_light.get('cost', '0'))

            speaker = elec.get('speaker', {})
            self.hallway_elec_speaker_check.set(speaker.get('check', False))
            self.hallway_elec_speaker_count.set(speaker.get('count', '0'))
            self.hallway_elec_speaker_cost.set(speaker.get('cost', '0'))

            outlet = elec.get('outlet', {})
            self.hallway_elec_outlet_check.set(outlet.get('check', False))
            self.hallway_elec_outlet_count.set(outlet.get('count', '0'))
            self.hallway_elec_outlet_cost.set(outlet.get('cost', '0'))

        # 通路の計算を更新
        if hasattr(self, 'update_hallway_arch_costs'):
            self.update_hallway_arch_costs()
        if hasattr(self, 'update_hallway_equipment_costs'):
            self.update_hallway_equipment_costs()
        if hasattr(self, 'update_hallway_subtotal'):
            self.update_hallway_subtotal()

        # 客室データ
        guest_room_data = data.get('guest_room_data', {})
        capacity = guest_room_data.get('capacity', {})
        self.capacity_japanese_room.set(capacity.get('japanese', '4.5'))
        self.capacity_japanese_western_room.set(capacity.get('japanese_western', '4.0'))
        self.capacity_japanese_bed_room.set(capacity.get('japanese_bed', '3.0'))
        self.capacity_western_room.set(capacity.get('western', '3.0'))

        arch = guest_room_data.get('arch', {})
        self.guest_arch_japanese_check.set(arch.get('japanese_check', False))
        self.guest_arch_japanese_area.set(arch.get('japanese_area', '0'))
        self.guest_arch_japanese_western_check.set(arch.get('japanese_western_check', False))
        self.guest_arch_japanese_western_area.set(arch.get('japanese_western_area', '0'))
        self.guest_arch_japanese_bed_check.set(arch.get('japanese_bed_check', False))
        self.guest_arch_japanese_bed_area.set(arch.get('japanese_bed_area', '0'))
        self.guest_arch_western_check.set(arch.get('western_check', False))
        self.guest_arch_western_area.set(arch.get('western_area', '0'))

        # 客室機械設備
        mech = guest_room_data.get('mech', {})
        self.guest_mech_ac_check.set(mech.get('ac_check', False))
        self.guest_mech_ac_count.set(mech.get('ac_count', '0'))
        self.guest_mech_wash_basin_check.set(mech.get('wash_basin_check', False))
        self.guest_mech_wash_basin_count.set(mech.get('wash_basin_count', '0'))
        self.guest_mech_sprinkler_check.set(mech.get('sprinkler_check', False))
        self.guest_mech_sprinkler_count.set(mech.get('sprinkler_count', '0'))

        # 客室電気設備
        elec = guest_room_data.get('elec', {})
        self.guest_elec_main_light_check.set(elec.get('main_light_check', False))
        self.guest_elec_main_light_count.set(elec.get('main_light_count', '0'))
        self.guest_elec_smoke_check.set(elec.get('smoke_check', False))
        self.guest_elec_smoke_count.set(elec.get('smoke_count', '0'))
        self.guest_elec_heat_detector_check.set(elec.get('heat_detector_check', False))
        self.guest_elec_heat_detector_count.set(elec.get('heat_detector_count', '0'))
        self.guest_elec_speaker_check.set(elec.get('speaker_check', False))
        self.guest_elec_speaker_count.set(elec.get('speaker_count', '0'))
        self.guest_elec_outlet_check.set(elec.get('outlet_check', False))
        self.guest_elec_outlet_count.set(elec.get('outlet_count', '0'))

        # 客室設備
        equipment = guest_room_data.get('equipment', {})
        self.guest_room_tv_cabinet.set(equipment.get('tv_cabinet', False))
        self.guest_room_headboard.set(equipment.get('headboard', False))
        self.guest_private_open_bath.set(equipment.get('private_open_bath', False))

        # 付加設備
        additional_equipment = data.get('additional_equipment', {})
        self.heat_source.set(additional_equipment.get('heat_source', False))
        self.elevator.set(additional_equipment.get('elevator', False))
        self.gas_equipment.set(additional_equipment.get('gas_equipment', False))
        self.septic_tank.set(additional_equipment.get('septic_tank', False))
        self.central_ac.set(additional_equipment.get('central_ac', False))
        self.fire_safety.set(additional_equipment.get('fire_safety', False))
        if hasattr(self, 'generator'):
            self.generator.set(additional_equipment.get('generator', False))

        # 計算を更新
        self.update_all_calculations()

    def delete_property(self):
        """選択された物件を削除"""
        selection = self.property_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "削除する物件を選択してください")
            return

        selected_text = self.property_listbox.get(selection[0])
        file, data = self.properties_data[selected_text]
        property_name = data.get('property_name', '名称未設定')

        result = messagebox.askyesno(
            "確認",
            f"物件「{property_name}」を削除しますか？\nこの操作は取り消せません。")

        if result:
            filepath = os.path.join(self.data_dir, file)
            try:
                os.remove(filepath)
                self.load_property_list()
                messagebox.showinfo("完了", f"物件「{property_name}」を削除しました")
            except Exception as e:
                messagebox.showerror("エラー", f"削除に失敗しました: {str(e)}")

    def on_tab_changed(self, event):
        """タブ切り替え時の処理（将来の拡張用）"""
        pass

    def on_closing(self):
        """ウィンドウを閉じる時の確認"""
        if self.data_modified:
            result = messagebox.askyesnocancel(
                "確認",
                "変更が保存されていません。保存しますか？")
            if result is None:  # キャンセル
                return
            elif result:  # はい
                self.save_current_property()

        self.root.destroy()

    def update_all_calculations(self):
        """
        建築初期設定（内装単価）の変更に伴い、関連する全ての計算を再実行する。
        """
        # 1. 温泉設備タブの建築項目の再計算
        self.update_onsen_calculations()
        self.update_wash_equipment_recommended()
        self.update_wash_equipment_costs()
        self.update_onsen_bath_arch_costs()
        self.update_onsen_subtotal()

        # レストランのタブに関する項目の再計算
        self.update_restaurant_arch_costs()
        self.update_restaurant_furniture_costs()
        self.update_restaurant_equipment_costs()
        self.update_restaurant_subtotal()

        # 2. その他の関連するタブの建築項目の再計算 (例: 客室、共用部など)
        # self.update_guest_room_arch_costs()
        # self.update_common_area_arch_costs()

        # 3. 必要に応じて、最終的な集計結果の再計算も行う
        # self.update_summary_costs()
        #self.update_cell_value()
        # ※ 必要に応じて、上記コメントアウト部分も実装・追加してください
        pass
    # 物件管理機能のタブに関する項目終了

    # コストサマリのタブに関する項目開始
    def check_area_warning(self, total_area):
        """
        合計面積が6000㎡以上になった場合に警告を表示

        Args:
            total_area (float): 現在の合計面積
        """
        if total_area >= self.area_warning_total_floor_area:
            # まだ警告を表示していない、または面積が閾値を超えて初めて表示する場合floor_num
            if not self.area_warning_shown:
                messagebox.showwarning(
                    "面積超過の警告",
                    f"【注意】\n\n"
                    f"階数ごとの面積合計が {self.area_warning_total_floor_area:,} ㎡を超えました。\n\n"
                    f"現在の合計面積: {total_area:,.1f} ㎡\n\n"
                    f"延べ床面積に対して\n"
                    f"スプリンクラー設備が必要になる可能性があります。"
                )
                self.area_warning_shown = True
        else:
            # 面積が閾値を下回った場合、フラグをリセット
            if self.area_warning_shown:
                self.area_warning_shown = False

    def calculate_date_difference(self, *args):
        """基準日と建物竣工年の日数差を計算"""
        try:

            base_str = self.base_date.get().strip()
            completion_str = self.completion_date.get().strip()

            if base_str and completion_str:
                base = datetime.strptime(base_str, '%Y/%m/%d')
                completion = datetime.strptime(completion_str, '%Y/%m/%d')
                diff = ( base-completion).days
                self.date_difference.set(diff)
                # 経過年数を計算してLCC分析タブの表示を更新
                elapsed_years = diff / 365.25  # うるう年を考慮
                self.update_lcc_year_marker(elapsed_years)
            else:
                self.date_difference.set(0)
                self.update_lcc_year_marker(0)
        except ValueError:
            # 日付フォーマットが不正な場合
            self.date_difference.set(0)
            self.update_lcc_year_marker(0)
        except Exception:
            self.date_difference.set(0)
            self.update_lcc_year_marker(0)

    def create_calculation_tab(self):
        """コストサマリタブの作成"""
        canvas = tk.Canvas(self.calc_frame)
        scrollbar = ttk.Scrollbar(self.calc_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        title_label = ttk.Label(main_frame, text="建物建設コスト算出", font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20), sticky=tk.W)

        row = 1

        ttk.Label(main_frame, text="建物の基本情報", font=sub_title_font).grid(row=row, column=0, columnspan=2,
                                                                                      pady=(20, 10), sticky=tk.W)
        row += 1

        ttk.Label(main_frame, text="建設地域:").grid(row=row, column=0, sticky=tk.W, pady=8)
        region_combo = ttk.Combobox(main_frame, textvariable=self.region, width=12)
        region_combo['values'] = ('東京都内', 'その他関東', '関西', '地方都市', '地方')
        region_combo.grid(row=row, column=1, sticky=tk.W, pady=8)
        region_combo.current(0)
        region_combo.bind('<<ComboboxSelected>>', self.on_region_changed)# 変更時に再計算
        row += 1

        ttk.Label(main_frame, text="基準日 (yyyy/mm/dd):").grid(row=row, column=0, sticky=tk.W, pady=8)
        base_date_entry = ttk.Entry(main_frame, textvariable=self.base_date, width=15)
        base_date_entry.grid(row=row, column=1, sticky=tk.W, pady=8)
        self.base_date.trace('w', self.calculate_date_difference)  # 変更時に日数差を計算
        row += 1

        ttk.Label(main_frame, text="建物竣工年 (yyyy/mm/dd):").grid(row=row, column=0, sticky=tk.W, pady=8)
        ttk.Entry(main_frame, textvariable=self.completion_date, width=15).grid(row=row, column=1, sticky=tk.W, pady=8)
        self.completion_date.trace('w', self.calculate_date_difference)  # 変更時に日数差を計算
        row += 1

        ttk.Label(main_frame, text="日数差:").grid(row=row, column=0, sticky=tk.W, pady=8)
        date_diff_label = ttk.Label(main_frame, textvariable=self.date_difference, font=("Arial", 10, "bold"))
        date_diff_label.grid(row=row, column=1, sticky=tk.W, pady=8)
        ttk.Label(main_frame, text="日").grid(row=row, column=1, sticky=tk.W, pady=8, padx=(60, 0))
        row += 1

        ttk.Label(main_frame, text="延床面積 (㎡):").grid(row=row, column=0, sticky=tk.W, pady=8)
        ttk.Entry(main_frame, textvariable=self.floor_area, width=15).grid(row=row, column=1, sticky=tk.W, pady=8)
        row += 1

        ttk.Label(main_frame, text="建物構造:").grid(row=row, column=0, sticky=tk.W, pady=8)
        structure_combo = ttk.Combobox(main_frame, textvariable=self.structure_type, width=12)
        structure_combo['values'] = ('木造', 'RC造', '鉄骨造', 'SRC造')
        structure_combo.grid(row=row, column=1, sticky=tk.W, pady=8)
        structure_combo.current(1)#デフォルトは　1:RC造
        structure_combo.bind('<<ComboboxSelected>>', self.update_construction_unit_price)

        row += 1

        ttk.Label(main_frame, text="階数:").grid(row=row, column=0, sticky=tk.W, pady=8)
        floors_frame = ttk.Frame(main_frame)
        floors_frame.grid(row=row, column=1, sticky=tk.W, pady=8)

        ttk.Label(floors_frame, text="地上").grid(row=0, column=1, sticky=tk.E, pady=2, padx=(0, 5))
        ttk.Entry(floors_frame, textvariable=self.floors, width=8).grid(row=0, column=2, sticky=tk.W, pady=2)
        ttk.Label(floors_frame, text="階").grid(row=0, column=3, sticky=tk.W, pady=2, padx=(5, 0))

        ttk.Label(floors_frame, text=" / ", font=('Arial', 10)).grid(row=0, column=4, sticky=tk.W, padx=(1, 1))

        ttk.Label(floors_frame, text="地下").grid(row=0, column=5, sticky=tk.E, pady=2, padx=(0, 5))
        ttk.Entry(floors_frame, textvariable=self.underground_floors, width=8).grid(row=0, column=6, sticky=tk.W,pady=2)
        ttk.Label(floors_frame, text="階").grid(row=0, column=7, sticky=tk.W, pady=2, padx=(5, 0))

        row += 1

        # ★★★ ここに階ごとの面積入力フレームを配置 ★★★
        self.floor_area_frame = ttk.Frame(main_frame)
        self.floor_area_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        row += 1

        # 建物断面図表示用のCanvas
        self.building_canvas = tk.Canvas(main_frame, width=300, height=400, bg='white', relief=tk.SUNKEN, borderwidth=2)
        self.building_canvas.grid(row=1, column=2, rowspan=15, padx=(20, 0), sticky=tk.N)
        # 階数変更時に断面イメージと階ごとの面積入力を更新
        self.floors.trace('w', self.update_building_image)
        self.underground_floors.trace('w', self.update_building_image)
        self.floors.trace('w', self.update_floor_area_inputs)
        self.underground_floors.trace('w', self.update_floor_area_inputs)

        # 客室数 (種類別):
        ttk.Label(main_frame, text="客室数 (種類別):").grid(row=row, column=0, sticky=tk.W, pady=8)
        room_detail_frame = ttk.Frame(main_frame)
        room_detail_frame.grid(row=row, column=1, sticky=tk.W, pady=8)

        ttk.Label(room_detail_frame, text="和室(10-15畳):").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.japanese_room_count = tk.StringVar(value="0")
        japanese_entry = ttk.Entry(room_detail_frame, textvariable=self.japanese_room_count, width=8,justify="right")
        japanese_entry.grid(row=0, column=1, sticky=tk.W, pady=2, padx=(5, 10))
        self.japanese_capacity_label = ttk.Label(room_detail_frame, textvariable=self.capacity_japanese_room,
                                                 text="人/室")
        self.japanese_capacity_label.grid(row=0, column=2, sticky=tk.W, pady=2)
        ttk.Label(room_detail_frame, text="人/室").grid(row=0, column=3, sticky=tk.W, pady=2)

        ttk.Label(room_detail_frame, text="和室・洋室:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.japanese_western_room_count = tk.StringVar(value="0")
        japanese_western_entry = ttk.Entry(room_detail_frame, textvariable=self.japanese_western_room_count, width=8,justify="right")
        japanese_western_entry.grid(row=1, column=1, sticky=tk.W, pady=2, padx=(5, 10))
        self.japanese_western_capacity_label = ttk.Label(room_detail_frame,
                                                         textvariable=self.capacity_japanese_western_room)
        self.japanese_western_capacity_label.grid(row=1, column=2, sticky=tk.W, pady=2)
        ttk.Label(room_detail_frame, text="人/室").grid(row=1, column=3, sticky=tk.W, pady=2)

        ttk.Label(room_detail_frame, text="和ベッド:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.japanese_bed_room_count = tk.StringVar(value="0")
        japanese_bed_entry = ttk.Entry(room_detail_frame, textvariable=self.japanese_bed_room_count, width=8,justify="right")
        japanese_bed_entry.grid(row=2, column=1, sticky=tk.W, pady=2, padx=(5, 10))
        self.japanese_bed_capacity_label = ttk.Label(room_detail_frame, textvariable=self.capacity_japanese_bed_room)
        self.japanese_bed_capacity_label.grid(row=2, column=2, sticky=tk.W, pady=2)
        ttk.Label(room_detail_frame, text="人/室").grid(row=2, column=3, sticky=tk.W, pady=2)

        ttk.Label(room_detail_frame, text="洋室:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.western_room_count = tk.StringVar(value="0")
        western_entry = ttk.Entry(room_detail_frame, textvariable=self.western_room_count, width=8,justify="right")
        western_entry.grid(row=3, column=1, sticky=tk.W, pady=2, padx=(5, 10))
        self.western_capacity_label = ttk.Label(room_detail_frame, textvariable=self.capacity_western_room)
        self.western_capacity_label.grid(row=3, column=2, sticky=tk.W, pady=2)
        ttk.Label(room_detail_frame, text="人/室").grid(row=3, column=3, sticky=tk.W, pady=2)

        self.total_guests_label = ttk.Label(room_detail_frame, text="合計宿泊人数: 0 人", font=("Arial", 10, "bold"),
                                            foreground="red")
        self.total_guests_label.grid(row=4, column=0, columnspan=4, sticky=tk.W, pady=(10, 5))

        self.japanese_room_count.trace('w', self.update_total_guests_realtime)
        self.japanese_western_room_count.trace('w', self.update_total_guests_realtime)
        self.japanese_bed_room_count.trace('w', self.update_total_guests_realtime)
        self.western_room_count.trace('w', self.update_total_guests_realtime)

        row += 1

        ttk.Label(main_frame, text="仕様グレード:").grid(row=row, column=0, sticky=tk.W, pady=8)
        spec_listbox_frame = ttk.Frame(main_frame)
        spec_listbox_frame.grid(row=row, column=1, sticky=tk.W, pady=8)

        spec_listbox = tk.Listbox(spec_listbox_frame, height=4, width=20, exportselection=False)
        spec_listbox.insert(0, "Gensen")
        spec_listbox.insert(1, "Premium 1")
        spec_listbox.insert(2, "Premium 2")
        spec_listbox.insert(3, "TAOYA 1")
        spec_listbox.insert(4, "TAOYA 2")
        spec_listbox.selection_set(0)
        spec_listbox.pack(side=tk.LEFT)

        spec_listbox.bind('<<ListboxSelect>>', self.update_selected_spec_grade)

        spec_listbox.bind('<<ListboxSelect>>', self.update_construction_unit_price, add='+')
        spec_listbox.bind('<<ListboxSelect>>', lambda e: self.update_lounge_arch_costs(), add='+')
        spec_listbox.bind('<<ListboxSelect>>', lambda e: self.update_amusement_arch_costs(), add='+')

        spec_scrollbar = ttk.Scrollbar(spec_listbox_frame, orient=tk.VERTICAL, command=spec_listbox.yview)
        spec_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        spec_listbox.configure(yscrollcommand=spec_scrollbar.set)

        #各小計フレーム表示の列幅設定　〇〇〇_subtotal_frame.grid_columnconfigure　
        w1=1
        w2=1
        w3=2
        w4=0

        self.spec_listbox = spec_listbox
        row += 1

        ttk.Label(main_frame, text="その他設備情報", font=sub_title_font).grid(row=row, column=0, columnspan=2,
                                                                                pady=(30, 10), sticky=tk.W)
        row += 1

        # 列幅の設定（すべてのその他設備情報フレームで共通）
        col_config = {
            0: 150,  # チェックボックス列
            1: 100,  # 金額列
            2: 30,  # "円"列
            3: 400  # 計算過程列
        }

        # ①熱源機器（昇降機設備の前に配置）
        heat_source_frame = ttk.Frame(main_frame)
        heat_source_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        # 列幅設定を他の設備と統一
        for col, width in col_config.items():
            heat_source_frame.grid_columnconfigure(col, minsize=width)

        # 1行目：チェックボックスと金額（他の設備と同じレイアウト）
        ttk.Checkbutton(heat_source_frame, text="熱源機器", variable=self.heat_source).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(heat_source_frame, textvariable=self.heat_source_cost, width=12, state='readonly',
                  justify=tk.RIGHT, font=("Arial", 9)).grid(row=0, column=1, sticky=tk.E, padx=(5, 5))
        ttk.Label(heat_source_frame, text="円", font=("Arial", 9)).grid(row=0, column=2, sticky=tk.E, padx=(5, 10))

        # 2行目：燃料種別と台数の入力（インデント付き）
        input_frame = ttk.Frame(heat_source_frame)
        input_frame.grid(row=1, column=0, columnspan=4, sticky=tk.W, pady=(5, 0))

        ttk.Label(input_frame, text="燃料種別:", font=("Arial", 8)).pack(side=tk.LEFT, padx=(20, 5))
        fuel_combo = ttk.Combobox(input_frame, textvariable=self.heat_source_fuel_type,
                                  values=("重油", "灯油", "ガス"), width=8, state='readonly')
        fuel_combo.pack(side=tk.LEFT, padx=(0, 15))
        fuel_combo.current(2)  # デフォルトは「ガス」

        ttk.Label(input_frame, text="台数:", font=("Arial", 8)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(input_frame, textvariable=self.heat_source_unit_count, width=5,
                  justify=tk.RIGHT).pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(input_frame, text="必要出力:", font=("Arial", 8)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(input_frame, textvariable=self.heat_source_required_output,
                  font=("Arial", 8), foreground="red").pack(side=tk.LEFT)

        # 3行目：計算過程表示（column=3で右側に配置、他の設備と統一）
        ttk.Label(heat_source_frame, textvariable=self.heat_source_process, font=("Arial", 8),
                  foreground="blue").grid(row=0, column=3, sticky=tk.W, padx=(10, 0))

        # 燃料種別と台数が変更されたときに再計算
        self.heat_source_fuel_type.trace('w', self.update_heat_source_equipment)
        self.heat_source_unit_count.trace('w', self.update_heat_source_equipment)

        row += 1

        # ①昇降機設備
        elevator_frame = ttk.Frame(main_frame)
        elevator_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        for col, width in col_config.items():
            elevator_frame.grid_columnconfigure(col, minsize=width)

        ttk.Checkbutton(elevator_frame, text="昇降機設備", variable=self.elevator).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(elevator_frame, textvariable=self.elevator_cost, width=12, state='readonly',
                  justify=tk.RIGHT, font=("Arial", 9)).grid(row=0, column=1, sticky=tk.E, padx=(5, 5))
        ttk.Label(elevator_frame, text="円", font=("Arial", 9)).grid(row=0, column=2, sticky=tk.E, padx=(5, 10))
        ttk.Label(elevator_frame, textvariable=self.elevator_process, font=("Arial", 8),
                  foreground="blue").grid(row=0, column=3, sticky=tk.W, padx=(10, 0))
        row += 1

        # ②ガス設備
        gas_frame = ttk.Frame(main_frame)
        gas_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        for col, width in col_config.items():
            gas_frame.grid_columnconfigure(col, minsize=width)

        ttk.Checkbutton(gas_frame, text="ガス設備", variable=self.gas_equipment).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(gas_frame, textvariable=self.gas_equipment_cost, width=12, state='readonly',
                  justify=tk.RIGHT, font=("Arial", 9)).grid(row=0, column=1, sticky=tk.E, padx=(5, 5))
        ttk.Label(gas_frame, text="円", font=("Arial", 9)).grid(row=0, column=2, sticky=tk.E, padx=(5, 10))
        ttk.Label(gas_frame, textvariable=self.gas_equipment_process, font=("Arial", 8),
                  foreground="blue").grid(row=0, column=3, sticky=tk.W, padx=(10, 0))
        row += 1

        # ③浄化槽設備
        septic_frame = ttk.Frame(main_frame)
        septic_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        for col, width in col_config.items():
            septic_frame.grid_columnconfigure(col, minsize=width)

        ttk.Checkbutton(septic_frame, text="浄化槽設備", variable=self.septic_tank).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(septic_frame, textvariable=self.septic_tank_cost, width=12, state='readonly',
                  justify=tk.RIGHT, font=("Arial", 9)).grid(row=0, column=1, sticky=tk.E, padx=(5, 5))
        ttk.Label(septic_frame, text="円", font=("Arial", 9)).grid(row=0, column=2, sticky=tk.E, padx=(5, 10))
        ttk.Label(septic_frame, textvariable=self.septic_tank_process, font=("Arial", 8),
                  foreground="blue").grid(row=0, column=3, sticky=tk.W, padx=(10, 0))
        row += 1

        # ④油送設備
        oil_frame = ttk.Frame(main_frame)
        oil_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        for col, width in col_config.items():
            oil_frame.grid_columnconfigure(col, minsize=width)

        ttk.Checkbutton(oil_frame, text="油送設備", variable=self.central_ac).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(oil_frame, textvariable=self.oil_supply_cost, width=12, state='readonly',
                  justify=tk.RIGHT, font=("Arial", 9)).grid(row=0, column=1, sticky=tk.E, padx=(5, 5))
        ttk.Label(oil_frame, text="円", font=("Arial", 9)).grid(row=0, column=2, sticky=tk.E, padx=(5, 10))
        ttk.Label(oil_frame, textvariable=self.oil_supply_process, font=("Arial", 8),
                  foreground="blue").grid(row=0, column=3, sticky=tk.W, padx=(10, 0))
        row += 1

        # ⑤受変電設備
        substation_frame = ttk.Frame(main_frame)
        substation_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        for col, width in col_config.items():
            substation_frame.grid_columnconfigure(col, minsize=width)

        ttk.Checkbutton(substation_frame, text="受変電設備", variable=self.fire_safety).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(substation_frame, textvariable=self.substation_cost, width=12, state='readonly',
                  justify=tk.RIGHT, font=("Arial", 9)).grid(row=0, column=1, sticky=tk.E, padx=(5, 5))
        ttk.Label(substation_frame, text="円", font=("Arial", 9)).grid(row=0, column=2, sticky=tk.E, padx=(5, 10))
        ttk.Label(substation_frame, textvariable=self.substation_process, font=("Arial", 8),
                  foreground="blue").grid(row=0, column=3, sticky=tk.W, padx=(10, 0))
        row += 1

        # ⑥自家発電設備
        generator_frame = ttk.Frame(main_frame)
        generator_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        for col, width in col_config.items():
            generator_frame.grid_columnconfigure(col, minsize=width)

        ttk.Checkbutton(generator_frame, text="自家発電設備", variable=self.generator).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(generator_frame, textvariable=self.generator_cost, width=12, state='readonly',
                  justify=tk.RIGHT, font=("Arial", 9)).grid(row=0, column=1, sticky=tk.E, padx=(5, 5))
        ttk.Label(generator_frame, text="円", font=("Arial", 9)).grid(row=0, column=2, sticky=tk.E, padx=(5, 10))
        ttk.Label(generator_frame, textvariable=self.generator_process, font=("Arial", 8),
                  foreground="blue").grid(row=0, column=3, sticky=tk.W, padx=(10, 0))
        row += 1

        calculate_btn = ttk.Button(main_frame, text="コスト計算", command=self.calculate_cost)
        calculate_btn.grid(row=row, column=0, columnspan=2, pady=30)
        row += 1

        # 温泉設備小計の表示
        onsen_subtotal_frame = ttk.Frame(main_frame)
        onsen_subtotal_frame.grid(row=row, column=0, columnspan=2, pady=(0, 10), sticky=(tk.W, tk.E))

        # 列の設定: 1列目がスペーサー（これは維持）
        onsen_subtotal_frame.grid_columnconfigure(0, weight=w1)
        onsen_subtotal_frame.grid_columnconfigure(1, weight=w2)
        onsen_subtotal_frame.grid_columnconfigure(2, weight=w3)
        onsen_subtotal_frame.grid_columnconfigure(3, weight=w4)

        # 項目名 (左寄せ)
        ttk.Label(onsen_subtotal_frame, text="温泉設備小計:", font=sheet_font).grid(row=0, column=0, sticky=tk.W)

        # 金額表示 (右寄せ)
        ttk.Label(onsen_subtotal_frame, textvariable=self.onsen_subtotal, font=sheet_font,
                  foreground="black").grid(row=0, column=2, sticky=tk.E, padx=(0, 5))

        # 単位 (右寄せ)
        ttk.Label(onsen_subtotal_frame, text=" 円", font=sheet_font).grid(row=0, column=3, sticky=tk.E)

        row += 1

        # レストラン小計の表示
        restaurant_subtotal_frame = ttk.Frame(main_frame)
        restaurant_subtotal_frame.grid(row=row, column=0, columnspan=2, pady=(0, 10), sticky=(tk.W, tk.E))

        # 列の設定: 1列目がスペーサー（これは維持）
        restaurant_subtotal_frame.grid_columnconfigure(0, weight=w1)
        restaurant_subtotal_frame.grid_columnconfigure(1, weight=w2)
        restaurant_subtotal_frame.grid_columnconfigure(2, weight=w3)
        restaurant_subtotal_frame.grid_columnconfigure(3, weight=w4)

        # 項目名 (左寄せ)
        ttk.Label(restaurant_subtotal_frame, text="レストラン小計:", font=sheet_font).grid(row=0, column=0,
                                                                                               sticky=tk.W)

        # 金額表示 (右寄せ)
        ttk.Label(restaurant_subtotal_frame, textvariable=self.restaurant_subtotal, font=sheet_font,
                  foreground="black").grid(row=0, column=2, sticky=tk.E, padx=(0, 5))

        # 単位 (右寄せ)
        ttk.Label(restaurant_subtotal_frame, text=" 円", font=sheet_font).grid(row=0, column=3, sticky=tk.E)

        row += 1

        # ラウンジ小計の表示
        lounge_subtotal_frame = ttk.Frame(main_frame)
        # 変更点: stickyから tk.E を削除し、tk.W のみ（左寄せ）にします
        lounge_subtotal_frame.grid(row=row, column=0, columnspan=2, pady=(0, 10), sticky=(tk.W, tk.E))

        # 列の設定: 1列目がスペーサー（これは維持）
        lounge_subtotal_frame.grid_columnconfigure(0, weight=w1)
        lounge_subtotal_frame.grid_columnconfigure(1, weight=w2)
        lounge_subtotal_frame.grid_columnconfigure(2, weight=w3)
        lounge_subtotal_frame.grid_columnconfigure(3, weight=w4)

        # 項目名 (左寄せ)
        ttk.Label(lounge_subtotal_frame, text="ラウンジ小計:", font=sheet_font).grid(row=0, column=0, sticky=tk.W)

        # 金額表示 (右寄せ)
        ttk.Label(lounge_subtotal_frame, textvariable=self.lounge_subtotal, font=sheet_font,
                  foreground="black").grid(row=0, column=2, sticky=tk.E, padx=(0, 5))

        # 単位 (右寄せ)
        ttk.Label(lounge_subtotal_frame, text=" 円", font=sheet_font).grid(row=0, column=3, sticky=tk.E)

        row += 1

        # 通路小計の表示
        hallway_subtotal_frame = ttk.Frame(main_frame)
        # 変更点: stickyから tk.E を削除し、tk.W のみ（左寄せ）にします
        hallway_subtotal_frame.grid(row=row, column=0, columnspan=2, pady=(0, 10), sticky=(tk.W, tk.E))

        # 列の設定: 1列目がスペーサー（これは維持）
        hallway_subtotal_frame.grid_columnconfigure(0, weight=w1)
        hallway_subtotal_frame.grid_columnconfigure(1, weight=w2)
        hallway_subtotal_frame.grid_columnconfigure(2, weight=w3)
        hallway_subtotal_frame.grid_columnconfigure(3, weight=w4)

        # 項目名 (左寄せ)
        ttk.Label(hallway_subtotal_frame, text="通路小計:", font=sheet_font).grid(row=0, column=0, sticky=tk.W)

        # 金額表示 (右寄せ)
        ttk.Label(hallway_subtotal_frame, textvariable=self.hallway_subtotal, font=sheet_font,
                  foreground="black").grid(row=0, column=2, sticky=tk.E, padx=(0, 5))

        # 単位 (右寄せ)
        ttk.Label(hallway_subtotal_frame, text=" 円", font=sheet_font).grid(row=0, column=3, sticky=tk.E)

        row += 1


        # アミューズメント小計の表示
        amusement_subtotal_frame = ttk.Frame(main_frame)
        # 変更点: stickyから tk.E を削除し、tk.W のみ（左寄せ）にします
        amusement_subtotal_frame.grid(row=row, column=0, columnspan=2, pady=(0, 9), sticky=(tk.W, tk.E))

        # 列の設定: 1列目がスペーサー（これは維持）
        amusement_subtotal_frame.grid_columnconfigure(0, weight=w1)
        amusement_subtotal_frame.grid_columnconfigure(1, weight=w2)
        amusement_subtotal_frame.grid_columnconfigure(2, weight=w3)
        amusement_subtotal_frame.grid_columnconfigure(3, weight=w4)

        # 項目名 (左寄せ)
        ttk.Label(amusement_subtotal_frame, text="アミューズメント小計:", font=sheet_font).grid(row=0, column=0,
                                                                                                    sticky=tk.W)
        # 金額表示 (右寄せ)
        ttk.Label(amusement_subtotal_frame, textvariable=self.amusement_subtotal, font=sheet_font,
                  foreground="black").grid(row=0, column=2, sticky=tk.E, padx=(0, 5))

        # 単位 (右寄せ)
        ttk.Label(amusement_subtotal_frame, text=" 円", font=sheet_font).grid(row=0, column=3, sticky=tk.E)

        row += 1

        # 客室小計の表示
        guest_subtotal_frame = ttk.Frame(main_frame)
        guest_subtotal_frame.grid(row=row, column=0, columnspan=2, pady=(0, 9), sticky=(tk.W, tk.E))

        guest_subtotal_frame.grid_columnconfigure(0, weight=w1)
        guest_subtotal_frame.grid_columnconfigure(1, weight=w2)
        guest_subtotal_frame.grid_columnconfigure(2, weight=w3)
        guest_subtotal_frame.grid_columnconfigure(3, weight=w4)

        ttk.Label(guest_subtotal_frame, text="客室小計:", font=sheet_font).grid(row=0, column=0, sticky=tk.W)
        ttk.Label(guest_subtotal_frame, textvariable=self.guest_subtotal, font=sheet_font,
                  foreground="black").grid(row=0, column=2, sticky=tk.E, padx=(0, 5))
        ttk.Label(guest_subtotal_frame, text=" 円", font=sheet_font).grid(row=0, column=3, sticky=tk.E)

        row += 1

        additional_subtotal_frame = ttk.Frame(main_frame)
        additional_subtotal_frame.grid(row=row, column=0, columnspan=2, pady=(0, 9), sticky=(tk.W, tk.E))

        additional_subtotal_frame.grid_columnconfigure(0, weight=w1)
        additional_subtotal_frame.grid_columnconfigure(1, weight=w2)
        additional_subtotal_frame.grid_columnconfigure(2, weight=w3)
        additional_subtotal_frame.grid_columnconfigure(3, weight=w4)

        ttk.Label(additional_subtotal_frame, text="付加設備小計:", font=sheet_font).grid(row=0, column=0, sticky=tk.W)
        ttk.Label(additional_subtotal_frame, textvariable=self.additional_equipment_subtotal, font=sheet_font,
                  foreground="black").grid(row=0, column=2, sticky=tk.E, padx=(0, 5))
        ttk.Label(additional_subtotal_frame, text=" 円", font=sheet_font).grid(row=0, column=3, sticky=tk.E)

        row += 1

        # 計算結果の表示
        cal_result_column_span=3

        ttk.Label(main_frame, text=f"計算結果 現在コラムスパンが {cal_result_column_span} となっている", font=sub_title_font).grid(row=row, column=0, columnspan=2,
                                                                                pady=(20, 5), sticky=tk.W)
        row += 1

        self.result_text = tk.Text(main_frame, height=15, width=80, wrap=tk.WORD)
        self.result_text.grid(row=row, column=0, columnspan=cal_result_column_span, pady=5, sticky=(tk.W, tk.E))#columnspan=2を変更するとまたがる列が広がる

        result_scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        result_scrollbar.grid(row=row, column=cal_result_column_span, sticky="ns", pady=5)
        self.result_text.configure(yscrollcommand=result_scrollbar.set)

        #建物断面図
        self.update_building_image()

    def get_floor_area_key(self, floor_num, is_underground=False):
        """階数と地下/地上を元に self.floor_areas のキーを生成"""
        if is_underground:
            return f"B{floor_num}F"
        else:
            return f"{floor_num}F"

    def update_building_image(self, *args):
        """建物断面イメージを更新（階ごとの面積に比例した幅で表示）"""
        try:
            above_floors = int(self.floors.get()) if self.floors.get() else 0
        except ValueError:
            above_floors = 0

        try:
            underground = int(self.underground_floors.get()) if self.underground_floors.get() else 0
        except ValueError:
            underground = 0

        # Canvasをクリア
        self.building_canvas.delete("all")

        # キャンバスのサイズを取得
        self.building_canvas.update_idletasks()
        canvas_width = self.building_canvas.winfo_width()
        canvas_height = self.building_canvas.winfo_height()

        # キャンバスがまだ描画されていない場合のデフォルト値
        if canvas_width <= 1:
            canvas_width = 300
        if canvas_height <= 1:
            canvas_height = 300

        # 描画パラメータ
        floor_height = 35  # 各階の高さ
        max_building_width = 200  # 最大幅
        text_offset_x = 8 # 階数テキストと建物外形との間のオフセット
        text_offset_y_to_bottom = 5

        total_floors = above_floors + underground

        if total_floors == 0:
            # 階数が0の場合は説明テキストを表示
            self.building_canvas.create_text(
                canvas_width // 2, canvas_height // 2,
                text="階数を入力してください",
                font=("Arial", 10),
                fill="gray",
                anchor='center'
            )
            return

        # 各階の面積を取得
        floor_areas_dict = {}
        max_area = 0

        # 地上階の面積を取得 (self.floor_areasから取得)
        for i in range(above_floors):
            floor_num = i + 1  # 1階から順に
            floor_key = f"{floor_num}F"
            area = 0

            if hasattr(self, 'floor_areas') and floor_key in self.floor_areas:
                try:
                    area_str = self.floor_areas[floor_key].get()
                    if area_str:
                        area = float(area_str)
                except (ValueError, AttributeError):
                    area = 0

            floor_areas_dict[floor_num] = area
            if area > max_area:
                max_area = area

        # 地下階の面積を取得 (self.floor_areasから取得)
        for i in range(underground):
            floor_num = i + 1  # B1, B2, B3...
            floor_key = f"B{floor_num}F"
            area = 0

            if hasattr(self, 'floor_areas') and floor_key in self.floor_areas:
                try:
                    area_str = self.floor_areas[floor_key].get()
                    if area_str:
                        area = float(area_str)
                except (ValueError, AttributeError):
                    area = 0

            floor_areas_dict[-floor_num] = area  # 負の値で保存（B1=-1, B2=-2...）
            if area > max_area:
                max_area = area

        # 最大面積が0の場合はデフォルト幅を使用
        if max_area == 0:
            max_area = 1  # 0除算を避けるため

        # 建物全体の高さを計算
        total_height = total_floors * floor_height

        # 描画開始位置を計算(中央配置)
        start_y = (canvas_height - total_height) // 2

        # 右端の位置を固定（キャンバス幅の中央より少し右）
        right_edge_x = canvas_width // 2 + max_building_width // 2


        # 地上階を描画(上から下へ: 最上階→...→2F→1F)
        for i in range(above_floors):
            # i=0が最上階、i=above_floors-1が1階
            floor_num = above_floors - i
            floor_y = start_y + i * floor_height

            # 面積に比例した幅を計算
            area = floor_areas_dict.get(floor_num, 0)
            if area > 0:
                floor_width = max_building_width * (area / max_area)
            else:
                floor_width = max_building_width * 0.3

            # 右端揃えのためのx座標（右端から逆算）
            start_x = right_edge_x - floor_width

            # 階の矩形
            self.building_canvas.create_rectangle(
                start_x, floor_y,
                start_x + floor_width, floor_y + floor_height,
                fill="" if (floor_num % 2) == 1 else "",
                outline="black",
                width=2
            )

            # 階のFL線
            self.building_canvas.create_line(
                right_edge_x, floor_y+ floor_height ,right_edge_x+20, floor_y+ floor_height,
                fill="black",
                width=1
            )

            # 階数表示（建物の矩形の右端の外側、各階の下端寄せ、「▽」付き）
            self.building_canvas.create_text(
                right_edge_x + text_offset_x,
                # Y座標を矩形の下端 (floor_y + floor_height) から上に text_offset_y_to_bottom だけオフセット
                floor_y + floor_height - text_offset_y_to_bottom,
                text=f"▽{floor_num}階",  # 「▽」を追加
                font=("Arial", 9, "bold"),
                fill="#696969",
                anchor='w'
            )

        # 地盤線(GL)を描画 - 1階の下端(地下階がある場合のみ)
        if underground > 0:
            ground_y = start_y + above_floors * floor_height
            self.building_canvas.create_line(
                right_edge_x - max_building_width - 20, ground_y,
                right_edge_x + 20, ground_y,
                fill="black",
                width=1
            )

            # 地盤線ラベル
            self.building_canvas.create_text(
                right_edge_x - max_building_width - 35, ground_y,
                text="GL",
                font=("Arial", 9, "bold"),
                fill="#696969"
            )

        # 地下階を描画(GLの下、上から下へ: B1F→B2F→B3F...)
        for i in range(underground):
            # i=0がB1F、i=1がB2F...
            floor_num = i + 1
            floor_y = start_y + above_floors * floor_height + i * floor_height

            # 面積に比例した幅を計算
            area = floor_areas_dict.get(-floor_num, 0)
            if area > 0:
                floor_width = max_building_width * (area / max_area)
            else:
                floor_width = max_building_width * 0.3

            # 右端揃えのためのx座標（右端から逆算）
            start_x = right_edge_x - floor_width

            # 階の矩形
            self.building_canvas.create_rectangle(
                start_x, floor_y,
                start_x + floor_width, floor_y + floor_height,
                fill="" if (floor_num % 2) == 1 else "",
                outline="black",
                width=2
            )

            # 階のFL線
            self.building_canvas.create_line(
                right_edge_x, floor_y + floor_height, right_edge_x + 20, floor_y + floor_height,
                fill="black",
                width=1
            )

            # 階数表示（建物の矩形の右端の外側、各階の下端寄せ、「▽」付き）
            self.building_canvas.create_text(
                right_edge_x + text_offset_x,
                # Y座標を矩形の下端 (floor_y + floor_height) から上に text_offset_y_to_bottom だけオフセット
                floor_y + floor_height - text_offset_y_to_bottom,
                text=f"▽B{floor_num}階",  # 「▽」を追加
                font=("Arial", 9, "bold"),
                fill="#696969",
                anchor='w'
            )

        # タイトル
        title_text = "建物断面図\n"
        if above_floors > 0 and underground > 0:
            title_text += f"(地上{above_floors}階 / 地下{underground}階)"
        elif above_floors > 0:
            title_text += f"(地上{above_floors}階)"
        else:
            title_text += f"(地下{underground}階)"

        self.building_canvas.create_text(
            canvas_width // 2, 30,
            text=title_text,
            font=("Arial", 9, "bold"),
            fill="black"
        )

    def get_total_floor_area(self):
        """全階の床面積の合計を計算して返す"""
        total_area = 0.0
        for floor_area_var in self.floor_areas.values():
            try:
                # カンマを除去してfloatに変換
                area = float(floor_area_var.get().replace(',', ''))
                total_area += area
            except (ValueError, AttributeError):
                # 入力がない、または不正な場合は0として扱う
                pass
        return total_area

    def update_total_floor_areas(self, *args):  # メソッド名をより汎用的に変更
        """延床面積を計算し、コストサマリとLCC分析タブの変数に設定する"""
        total_area = self.get_total_floor_area()

        # 1. LCC分析タブの建物面積 (数値として使用される変数) を更新
        self.lcc_building_area.set(f"{int(total_area):,}")
        self.update_construction_unit_price()

        # 2. コストサマリタブの合計面積表示 (表示用の文字列) を更新 ★★★
        self.cost_summary_total_area_var.set(f"合計: {int(total_area):,} ㎡")

        #################################チェック用に追加
        # 4階から10階までの面積をチェック
        for floor_num in range(4, 11):  # 4階から10階まで
            floor_key = f"{floor_num}F"
            try:
                if floor_key in self.floor_areas:
                    # self.floor_areas[floor_key] は 'StringVar' オブジェクトであることが判明
                    floor_string_var = self.floor_areas[floor_key]
                    floor_value = floor_string_var.get()  # 値の取得は .get() でOK

                    if floor_value:  # 空文字列でない場合のみ

                        # --- ★不正形式のバリデーションとクリア処理★ ---
                        # 例: "0333" のような、先頭が '0' で始まる複数桁の数字を検出
                        if floor_value.isdigit() and len(floor_value) > 1 and floor_value.startswith('0'):
                            # 1. エラーメッセージの表示（仮の出力）
                            print(f"エラー: {floor_num}階の面積の値'{floor_value}'は不正な形式です。（先頭の0は不可）")

                            # 2. 【追加/修正】メッセージボックスの表示
                            # messagebox.showerror(タイトル, メッセージ) を使用
                            messagebox.showerror(
                                "入力エラー",
                                f"{floor_num}階の面積: 値 '{floor_value}' は不正な形式です。\n"
                                "先頭に '0' を含む複数桁の数字は入力できません。"
                            )
                            # StringVar の場合は set("") を使います
                            floor_string_var.set("")

                            # 3. この階の処理をスキップ
                            continue
                            # ----------------------------------------------

                        # 値が有効な場合、処理を続行
                        floor_area = float(floor_value)
                        print(f"{floor_num}階の面積: {floor_area} ㎡")

                        if floor_area >= 1500:
                            print(f"{floor_num}階: SP必要 (1500㎡以上)")
                    else:
                        print(f"{floor_num}階の面積: 未入力")
                else:
                    pass  # 階が存在しない場合は何も出力しない
            except ValueError:
                print(f"エラー: {floor_num}階の値が無効です")
                # float()変換エラー時も値をクリアしたい場合は、この行の下に set("") を追加
                # floor_string_var.set("")
        ##################################

    def update_floor_area_inputs(self, *args):
        """階数に応じて各階の面積入力欄を動的に生成"""
        # 既存の入力欄をクリア
        for widget in self.floor_area_frame.winfo_children():
            widget.destroy()

        # 現在の階数を取得
        try:
            above_floors = int(self.floors.get()) if self.floors.get() else 0
        except ValueError:
            above_floors = 0

        try:
            underground = int(self.underground_floors.get()) if self.underground_floors.get() else 0
        except ValueError:
            underground = 0

        total_floors = above_floors + underground

        if total_floors == 0:
            return

        # ★★★ 現在必要な階のキーセットを作成 ★★★
        current_floor_keys = set()
        for i in range(above_floors):
            floor_num = above_floors - i
            current_floor_keys.add(f"{floor_num}F")
        for i in range(underground):
            floor_num = i + 1
            current_floor_keys.add(f"B{floor_num}F")

        # ★★★ 不要な階の面積データを削除 ★★★
        keys_to_remove = [key for key in self.floor_areas.keys() if key not in current_floor_keys]
        for key in keys_to_remove:
            del self.floor_areas[key]

        # タイトル
        title_label = ttk.Label(self.floor_area_frame, text="階ごとの面積入力 (㎡)",
                                font=("Arial", 9), foreground="black")
        title_label.grid(row=0, column=0, columnspan=4, pady=(5, 10), sticky=tk.W)

        current_row = 1

        # 地上階の入力欄を作成（上層階から順に）
        for i in range(above_floors):
            floor_num = above_floors - i
            floor_key = f"{floor_num}F"

            # まだ辞書に存在しない場合は新規作成
            if floor_key not in self.floor_areas:
                self.floor_areas[floor_key] = tk.StringVar(value="0")
                self.floor_areas[floor_key].trace_add('write', self.update_total_floor_areas)
                self.floor_areas[floor_key].trace_add('write', self.update_building_image)

            # ラベルと入力欄を配置
            ttk.Label(self.floor_area_frame, text=f"{floor_num}階:").grid(
                row=current_row, column=0, sticky=tk.W, padx=(10, 5), pady=3
            )
            ttk.Entry(self.floor_area_frame, textvariable=self.floor_areas[floor_key],
                      width=10, justify=tk.RIGHT).grid(
                row=current_row, column=1, sticky=tk.W, padx=(0, 5), pady=3
            )
            ttk.Label(self.floor_area_frame, text="㎡").grid(
                row=current_row, column=2, sticky=tk.W, padx=(0, 20), pady=3
            )

            current_row += 1

        # 地下階の入力欄を作成
        for i in range(underground):
            floor_num = i + 1
            floor_key = f"B{floor_num}F"

            # まだ辞書に存在しない場合は新規作成
            if floor_key not in self.floor_areas:
                self.floor_areas[floor_key] = tk.StringVar(value="0")

            # トレースを設定（新規・既存問わず毎回設定）
            self.floor_areas[floor_key].trace_add('write', self.update_total_floor_areas)
            self.floor_areas[floor_key].trace_add('write', self.update_building_image)

            # ラベルと入力欄を配置
            ttk.Label(self.floor_area_frame, text=f"地下{floor_num}階:").grid(
                row=current_row, column=0, sticky=tk.W, padx=(10, 5), pady=3
            )
            ttk.Entry(self.floor_area_frame, textvariable=self.floor_areas[floor_key],
                      width=10, justify=tk.RIGHT).grid(
                row=current_row, column=1, sticky=tk.W, padx=(0, 5), pady=3
            )
            ttk.Label(self.floor_area_frame, text="㎡").grid(
                row=current_row, column=2, sticky=tk.W, padx=(0, 20), pady=3
            )

            current_row += 1

        # 合計面積表示（オプション）

        ttk.Label(self.floor_area_frame, textvariable=self.cost_summary_total_area_var,  # <-- 変数名を修正
                  font=("Arial", 10, "bold"), foreground="black").grid(
            row=current_row, column=0, columnspan=3, sticky=tk.W, padx=(30, 0), pady=(10, 5)
        )

    def update_selected_spec_grade(self, event):
        """リストボックスで選択された仕様グレードを self.selected_spec_grade に設定する"""
        try:
            # 選択されているインデックスを取得
            selected_indices = event.widget.curselection()
            if selected_indices:
                # 選択されている項目名を取得
                selected_grade = event.widget.get(selected_indices[0])
                # 変数に設定
                self.selected_spec_grade.set(selected_grade)
            else:
                self.selected_spec_grade.set("未選択")  # 選択が解除された場合
        except Exception as e:
            print(f"仕様グレードの更新エラー: {e}")

    def create_arch_settings_tab(self):
        """建築初期設定タブの作成"""
        unit_cost_data_vars = self.arch_unit_costs
        base_cost_data = self.arch_base_cost_data

        canvas = tk.Canvas(self.arch_settings_frame)
        scrollbar = ttk.Scrollbar(self.arch_settings_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        row = 0
        ttk.Label(main_frame, text="建築設定", font=title_font).grid(row=row, column=0, sticky=tk.W,
                                                                     pady=(0, 20))

        row += 1

        # ★★★ 【新規追加】建設地域掛け率設定テーブルを作成 ★★★
        row = self.create_region_rate_table(main_frame, row)

        # ★★★ 【既存】建物構造別設定テーブルを作成 ★★★
        row = self.create_structure_unit_cost_table(main_frame, row)

        title_button_frame = ttk.Frame(main_frame)
        title_button_frame.grid(row=row, column=0, columnspan=17, pady=(20, 10), sticky=tk.W)

        row += 1

        ttk.Label(title_button_frame, text="内装単価設定 (円/㎡)", font=sub_title_font).pack(side=tk.LEFT)

        ttk.Button(title_button_frame, text="追加", command=self.add_arch_row).pack(side=tk.LEFT, padx=(10, 5))
        ttk.Button(title_button_frame, text="削除", command=self.delete_arch_row).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(title_button_frame, text="保存", command=self.save_arch_data).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(title_button_frame, text="リセット", command=self.reset_arch_data).pack(side=tk.LEFT)

        header_row = row

        ttk.Label(main_frame, text="室用途", font=("Arial", 10)).grid(row=header_row, column=0, sticky=tk.W, padx=5)
        ttk.Label(main_frame, text="天井高(m)", font=("Arial", 10)).grid(row=header_row, column=1, sticky=tk.W, padx=5)

        sub_header_row = header_row + 1

        header_titles = ["内装単価設定(床)", "内装単価設定(天井)", "内装単価設定(壁)"]
        grade_titles = ["Gensen", "Premium 1", "Premium 2", "TAOYA 1", "TAOYA 2"]
        col_offset = 2

        for i, title in enumerate(header_titles):
            main_header_padx = (5, 0) if i > 0 else (0, 0)
            label = ttk.Label(main_frame, text=title, font=("Arial", 10), anchor="center", relief=tk.FLAT)
            label.grid(row=header_row, column=col_offset, columnspan=len(grade_titles), sticky="ew",
                       padx=main_header_padx)

            for j, grade in enumerate(grade_titles):
                grade_padx_left = 3 + (5 if i > 0 and j == 0 else 0)
                ttk.Label(main_frame, text=grade, font=("Arial", 9)).grid(row=sub_header_row, column=col_offset + j,
                                                                          sticky="ew", padx=(grade_padx_left, 3))
            col_offset += len(grade_titles)

        row = sub_header_row + 1

        for data_row_vars in unit_cost_data_vars:
            name, height_var, cost_vars = data_row_vars

            ttk.Label(main_frame, text=name, font=("Arial", 9)).grid(row=row, column=0, sticky=tk.W, padx=5, pady=1)

            height_entry = ttk.Entry(main_frame, textvariable=height_var, width=5, justify=tk.RIGHT)
            height_entry.grid(row=row, column=1, sticky="ew", padx=1, pady=1)

            for i, var in enumerate(cost_vars):
                col_index = i + 2
                extra_padx_left = 5 if i == 5 or i == 10 else 0
                cost_entry = ttk.Entry(main_frame, textvariable=var, width=6, justify=tk.RIGHT)
                cost_entry.grid(row=row, column=col_index, sticky="ew", padx=(1 + extra_padx_left, 1), pady=1)

            row += 1

        row += 2
        ttk.Label(main_frame, text="下地単価設定 (Ver1.0では未設定)", font=sub_title_font).grid(row=row,
                                                                                                column=0,
                                                                                                columnspan=17,
                                                                                                pady=(20, 10),
                                                                                                sticky=tk.W)
        row += 1

        sub_frame = ttk.Frame(main_frame)
        sub_frame.grid(row=row, column=0, columnspan=17, sticky=tk.W)

        items_per_col = 10

        for i, item in enumerate(base_cost_data):
            list_row = i % items_per_col
            list_col = i // items_per_col

            ttk.Label(sub_frame, text=f"{i + 1}. {item}").grid(row=list_row, column=list_col * 2, sticky=tk.W,
                                                               padx=(0, 5))
            ttk.Label(sub_frame, text="未設定").grid(row=list_row, column=list_col * 2 + 1, sticky=tk.W, padx=(0, 20))

    def create_region_rate_table(self, parent_frame, start_row):
        """建設地域掛け率設定の表を作成する"""
        current_row = start_row

        # セクションタイトル
        ttk.Label(parent_frame, text="建設地域掛け率設定",
                  font=sub_title_font).grid(
            row=current_row, column=0, columnspan=6, pady=(15, 5), sticky=tk.W
        )
        current_row += 1

        # 説明文
        ttk.Label(parent_frame,
                  text="※ コストサマリタブで選択する建設地域に応じた再調達価格の掛け率を設定します",
                  font=("Arial", 9), foreground="gray").grid(
            row=current_row, column=0, columnspan=6, pady=(0, 10), sticky=tk.W
        )
        current_row += 1

        # ヘッダー
        ttk.Label(parent_frame, text="建設地域", font=sheet_font).grid(
            row=current_row, column=0, sticky=tk.W, padx=5
        )
        ttk.Label(parent_frame, text="掛け率 (%)", font=sheet_font).grid(
            row=current_row, column=1, sticky=tk.W, padx=5
        )
        current_row += 1

        # データ行
        region_order = ['東京都内', 'その他関東', '関西', '地方都市', '地方']

        for region in region_order:
            # 地域名
            ttk.Label(parent_frame, text=region, font=sheet_font).grid(
                row=current_row, column=0, sticky=tk.W, padx=5
            )

            # 掛け率入力欄
            ttk.Entry(parent_frame,
                      textvariable=self.region_rates[region],
                      width=10,
                      justify=tk.RIGHT,
                      font=sheet_font).grid(
                row=current_row, column=1, sticky=(tk.W, tk.E), padx=2, pady=1
            )

            current_row += 1

        # 保存・リセットボタン
        button_frame = ttk.Frame(parent_frame)
        button_frame.grid(row=current_row, column=0, columnspan=6,
                          pady=(10, 20), sticky=tk.W)

        ttk.Button(button_frame, text="地域掛け率を保存",
                   command=self.save_region_rate_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="地域掛け率をリセット",
                   command=self.reset_region_rate_data).pack(side=tk.LEFT, padx=5)

        current_row += 1

        return current_row

    def save_region_rate_data(self):
        """建設地域掛け率データをJSONファイルに保存"""
        data_to_save = {}

        for region, var in self.region_rates.items():
            try:
                rate = float(var.get())
                data_to_save[region] = rate
            except ValueError:
                messagebox.showwarning(
                    "入力エラー",
                    f"{region}の掛け率が正しく入力されていません。"
                )
                return

        try:
            with open(self.region_rate_file, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("保存完了", "建設地域掛け率データが保存されました。")

            # 掛け率設定を保存した時に再計算
            self.update_all_calculations()

        except Exception as e:
            messagebox.showerror("保存エラー",
                                 f"データの保存に失敗しました:\n{str(e)}")

    def load_region_rate_data(self):
        """建設地域掛け率データをJSONファイルから読み込む"""
        if os.path.exists(self.region_rate_file):
            try:
                with open(self.region_rate_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return data
            except Exception as e:
                messagebox.showwarning(
                    "読み込みエラー",
                    f"データの読み込みに失敗しました:\n{str(e)}\n"
                    "初期データを使用します。"
                )
                return None
        return None

    def reset_region_rate_data(self):
        """建設地域掛け率データを初期状態にリセット"""
        result = messagebox.askyesno(
            "確認",
            "建設地域掛け率データを初期状態にリセットしますか?"
        )

        if result:
            # デフォルト値に戻す
            default_rates = {
                '東京都内': "100",
                'その他関東': "90",
                '関西': "85",
                '地方都市': "80",
                '地方': "75"
            }

            for region, default_value in default_rates.items():
                self.region_rates[region].set(default_value)

            # ファイルを削除
            if os.path.exists(self.region_rate_file):
                try:
                    os.remove(self.region_rate_file)
                except:
                    pass

            messagebox.showinfo("リセット完了",
                                "建設地域掛け率データが初期状態にリセットされました。")

            # 再計算を実行
            self.update_all_calculations()

    def get_region_rate(self):
        """現在選択されている建設地域の掛け率を取得する"""
        try:
            region = self.region.get()
            if region in self.region_rates:
                rate = float(self.region_rates[region].get())
                return rate / 100.0  # パーセントから小数に変換
            else:
                return 1.0  # デフォルト値
        except ValueError:
            return 1.0  # エラー時はデフォルト値

    def on_region_changed(self, event=None):
        """
        建設地域が変更されたときに呼ばれる関数
        再調達価格を再計算する
        """
        try:
            print(f"建設地域が変更されました: {self.region.get()}")

            # 再調達価格を再計算
            self.update_construction_unit_price()

            # LCCタブが初期化されている場合は、LCC関連の計算も更新
            if hasattr(self, 'update_all_calculations'):
                self.update_all_calculations()

            print("地域変更に伴う再計算が完了しました")

        except Exception as e:
            print(f"地域変更時のエラー: {e}")

    def calculate_cost(self):
        """コスト計算のロジック"""
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "この機能はまだ実装されていません。")

    def update_settings_display(self, *args):
        """設定タブの表示を更新"""
        self.misc_rate_display.config(text=f"現在の諸経費率: {self.misc_cost_rate.get():.2f}")
        self.spec1_display.config(text=f"仕様1: {self.spec1_rate.get():.2f} ({(self.spec1_rate.get() * 100):.0f}%)")
        self.spec2_display.config(text=f"仕様2: {self.spec2_rate.get():.2f} ({(self.spec2_rate.get() * 100):.0f}%)")
        self.spec3_display.config(text=f"仕様3: {self.spec3_rate.get():.2f} ({(self.spec3_rate.get() * 100):.0f}%)")

    def reset_settings(self):
        """設定値を初期値にリセット"""
        self.misc_cost_rate.set(0.15)
        self.spec1_rate.set(0.10)
        self.spec2_rate.set(0.25)
        self.spec3_rate.set(0.45)
        messagebox.showinfo("リセット完了", "すべての設定が初期値にリセットされました。")
        self.update_settings_display()

    def show_current_settings(self):
        """現在の設定内容を表示"""
        current_settings = (
            f"--- 現在の設定 --- \n"
            f"諸経費率: {self.misc_cost_rate.get():.2f} ({self.misc_cost_rate.get() * 100:.0f}%) \n"
            f"仕様1割増率: {self.spec1_rate.get():.2f} ({self.spec1_rate.get() * 100:.0f}%) \n"
            f"仕様2割増率: {self.spec2_rate.get():.2f} ({self.spec2_rate.get() * 100:.0f}%) \n"
            f"仕様3割増率: {self.spec3_rate.get():.2f} ({self.spec3_rate.get() * 100:.0f}%)"
        )
        messagebox.showinfo("現在の設定", current_settings)

    def update_additional_equipment_costs(self, *args):
        """付加設備の金額と計算過程を更新"""
        try:
            # 客室数の合計を取得
            total_rooms = 0
            total_rooms += int(self.japanese_room_count.get()) if self.japanese_room_count.get() else 0
            total_rooms += int(
                self.japanese_western_room_count.get()) if self.japanese_western_room_count.get() else 0
            total_rooms += int(self.japanese_bed_room_count.get()) if self.japanese_bed_room_count.get() else 0
            total_rooms += int(self.western_room_count.get()) if self.western_room_count.get() else 0

            # 合計宿泊人数を取得
            total_guests = self.get_total_guests()

            # ①熱源設備
            self.update_heat_source_equipment()
            # ①昇降機設備
            if total_rooms > 0:
                elevator_cost = total_rooms * (425000+30000)
                self.elevator_cost.set(f"{int(elevator_cost):,}")
                self.elevator_process.set(f"計算式: {total_rooms}室 × [425,000(ELV)+30,000(DW)] = {int(elevator_cost):,}円")
            else:
                self.elevator_cost.set("0")
                self.elevator_process.set("客室数を入力してください")

            #  ②ガス設備
            if total_guests > 0:
                gas_cost = total_guests * 10000 + 2000000
                self.gas_equipment_cost.set(f"{int(gas_cost):,}")
                self.gas_equipment_process.set(
                    f"計算式: {int(total_guests)}人 × 10,000 + 2,000,000 = {int(gas_cost):,}円(厨房利用程度0.75m3/人)")
            else:
                self.gas_equipment_cost.set("0")
                self.gas_equipment_process.set("宿泊人数を入力してください")

            # ③浄化槽設備
            if total_guests > 0:
                septic_cost = total_guests * 300000 + 5000000
                self.septic_tank_cost.set(f"{int(septic_cost):,}")
                self.septic_tank_process.set(
                    f"計算式: {int(total_guests)}人 × 300,000 + 5,000,000 = {int(septic_cost):,}円")
            else:
                self.septic_tank_cost.set("0")
                self.septic_tank_process.set("宿泊人数を入力してください")

            # ④油送設備
            if total_guests > 0:
                oil_cost = total_guests * 10000 + 3000000
                self.oil_supply_cost.set(f"{int(oil_cost):,}")
                self.oil_supply_process.set(
                    f"計算式: {int(total_guests)}人 × 10,000 + 3,000,000 = {int(oil_cost):,}円")
            else:
                self.oil_supply_cost.set("0")
                self.oil_supply_process.set("宿泊人数を入力してください")

            # ⑤受変電設備
            if total_guests > 0:
                substation_cost = total_guests * 500000 + 1000000
                self.substation_cost.set(f"{int(substation_cost):,}")
                self.substation_process.set(
                    f"計算式: {int(total_guests)}人 × 500,000 + 1,000,000 = {int(substation_cost):,}円")
            else:
                self.substation_cost.set("0")
                self.substation_process.set("宿泊人数を入力してください")

            # ⑥自家発電設備
            if total_guests > 0:
                generator_cost = total_guests * 35000 + 1000000
                self.generator_cost.set(f"{int(generator_cost):,}")
                self.generator_process.set(
                    f"計算式: {int(total_guests)}人 × 35,000 + 1,000,000 = {int(generator_cost):,}円")
            else:
                self.generator_cost.set("0")
                self.generator_process.set("宿泊人数を入力してください")

        except (ValueError, AttributeError) as e:
            # エラー時は0を設定
            self.heat_source_cost.set("0")
            self.elevator_cost.set("0")
            self.gas_equipment_cost.set("0")
            self.septic_tank_cost.set("0")
            self.oil_supply_cost.set("0")
            self.substation_cost.set("0")
            self.generator_cost.set("0")

        # 小計を更新
        self.update_additional_equipment_subtotal()

    def update_heat_source_equipment(self, *args):
        """熱源機器の必要出力を計算し、適切な機器を選定して金額を表示"""
        try:
            # 合計宿泊人数を取得
            total_guests = self.get_total_guests()

            if total_guests <= 0:
                self.heat_source_cost.set("0")
                self.heat_source_required_output.set("0 Kw")
                self.heat_source_process.set("宿泊人数を入力してください")
                return

            # ①必要な出力を計算
            # 出力 = 0.0016 × 宿泊人数 × (((60+55)/2-5)×20×2-(60-55)×30) / 2
            output = 0.0016 * total_guests * (((60 + 55) / 2 - 5) * 20 * 2 - (60 - 55) * 30) / 2

            # 必要出力を表示
            self.heat_source_required_output.set(f"{output:.1f} Kw")

            # ②燃料種別を取得
            fuel_type = self.heat_source_fuel_type.get()

            # ③台数を取得
            try:
                unit_count = int(self.heat_source_unit_count.get())
                if unit_count <= 0:
                    unit_count = 1
            except (ValueError, AttributeError):
                unit_count = 1

            # 1台あたりの必要出力
            output_per_unit = output / unit_count

            # ④設備設定タブの熱源機器設定から適切な機器を選定
            selected_equipment = None
            selected_cost = 0
            selected_name = ""

            for data_vars in self.heat_source_unit_data:
                no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var, \
                    new_equipment_cost_var, area_ratio_var, fuel_type_var, output_var = data_vars

                # 燃料種別が一致するかチェック
                equipment_fuel_type = fuel_type_var.get()
                if equipment_fuel_type != fuel_type:
                    continue

                # 出力を取得（例: "349Kw" から 349 を抽出）
                equipment_output_str = output_var.get()
                try:
                    # "Kw" を削除して数値を取得
                    equipment_output = float(
                        equipment_output_str.replace("Kw", "").replace("kw", "").replace("KW", "").strip())

                    # 1台あたりの必要出力以上の機器を探す
                    if equipment_output >= output_per_unit:
                        # まだ選定されていないか、より小さい（適切な）出力の機器なら更新
                        if selected_equipment is None or equipment_output < selected_equipment:
                            selected_equipment = equipment_output
                            selected_name = abbrev_var.get()

                            # 機器単価を取得
                            try:
                                unit_cost = float(new_equipment_cost_var.get().replace(',', ''))
                                # 台数分の金額を計算
                                selected_cost = unit_cost * unit_count
                            except (ValueError, AttributeError):
                                selected_cost = 0

                            # 機器名を保存
                            selected_name = abbrev_var.get()

                except (ValueError, AttributeError):
                    continue

            # 結果を設定
            if selected_equipment is not None:
                self.heat_source_cost.set(f"{int(selected_cost):,}")
                self.heat_source_process.set(
                    f"必要出力: {output:.1f}Kw ÷ {unit_count}台 = {output_per_unit:.1f}Kw/台 → "
                    f"{selected_name}({selected_equipment:.1f}Kw) × {unit_count}台 = {int(selected_cost):,}円"
                )
            else:
                self.heat_source_cost.set("0")
                self.heat_source_process.set(
                    f"必要出力: {output:.1f}Kw ÷ {unit_count}台 = {output_per_unit:.1f}Kw/台 → "
                    f"適切な機器が見つかりません（燃料種別: {fuel_type}）"
                )

        except Exception as e:
            print(f"熱源機器計算エラー: {e}")
            import traceback
            traceback.print_exc()
            self.heat_source_cost.set("0")
            self.heat_source_required_output.set("0 Kw")
            self.heat_source_process.set("計算エラー")

    def update_additional_equipment_subtotal(self, *args):
        """付加設備小計を計算"""
        total = 0

        # ①熱源設備
        if self.heat_source.get():
            try:
                cost = float(self.heat_source_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass
        # ①昇降機設備
        if self.elevator.get():
            try:
                cost = float(self.elevator_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # ②ガス設備
        if self.gas_equipment.get():
            try:
                cost = float(self.gas_equipment_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass
        # ③浄化槽設備
        if self.septic_tank.get():
            try:
                cost = float(self.septic_tank_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

            # 浄化槽設備がチェックされている場合、再調達価格を更新
        try:

            # 基本の再調達価格を取得（建築基準単価×面積）
            area_str = self.lcc_building_area.get().replace(',', '').strip()
            unit_price_str = self.lcc_construction_unit_price.get().replace(',', '').strip()

            print(f"面積: {area_str}")
            print(f"建築基準単価: {unit_price_str}")

            if area_str and unit_price_str:
                area = float(area_str)
                unit_price = float(unit_price_str)
                base_reconstruction_cost = area * unit_price

                print(f"基本再調達価格: {base_reconstruction_cost:,.0f}円")

                # 浄化槽設備の金額を取得
                septic_cost = 0
                if self.septic_tank.get():
                    septic_cost_str = self.septic_tank_cost.get().replace(',', '').strip()
                    if septic_cost_str:
                        septic_cost = float(septic_cost_str)
                        print(f"浄化槽工事金額: {septic_cost:,.0f}円")
                else:
                    print(f"浄化槽チェックなし")

                # 新しい再調達価格 = 基本再調達価格 + 浄化槽工事金額
                new_reconstruction_cost = base_reconstruction_cost + septic_cost

                print(f"新しい再調達価格: {new_reconstruction_cost:,.0f}円")

                # 再調達価格を更新
                self.new_construction_cost.set(f"{int(new_reconstruction_cost):,}")

                # LCC修繕計画表の比率を更新
                self.update_lcc_ratios_for_septic_tank(base_reconstruction_cost, septic_cost)

        except (ValueError, AttributeError) as e:
            print(f"再調達価格更新エラー: {e}")
            import traceback
            traceback.print_exc()
        # ④油送設備
        if self.central_ac.get():
            try:
                cost = float(self.oil_supply_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # ⑤受変電設備
        if self.fire_safety.get():
            try:
                cost = float(self.substation_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # ⑥自家発電設備
        if self.generator.get():
            try:
                cost = float(self.generator_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # 小計を更新
        self.additional_equipment_subtotal.set(f"{int(total):,}")

    def update_lcc_ratios_for_septic_tank(self, base_cost, septic_cost):
        """
        浄化槽設備による LCC修繕計画表の比率更新

        Args:
            base_cost: 変更前の再調達価格（建築基準単価×面積）
            septic_cost: 浄化槽工事の金額
        """
        try:
            # LCC修繕計画表が初期化されていない場合は処理をスキップ
            if not hasattr(self, 'table_vars'):
                return

            # print(f"\n=== 浄化槽設備による比率更新 ===")
            # print(f"基本再調達価格: {base_cost:,.0f}円")
            # print(f"浄化槽工事金額: {septic_cost:,.0f}円")

            # 初期比率の定義
            DEFAULT_ARCH_RATIO = 40.0  # 建築工事計の初期比率
            DEFAULT_HVAC_RATIO = 12.0  # 空調設備工事計の初期比率
            DEFAULT_SANITARY_RATIO = 8.0  # 衛生設備計の初期比率
            DEFAULT_ELEC_RATIO = 15.5  # 電気設備計の初期比率
            DEFAULT_ELEVATOR_RATIO = 0.87  # 昇降機設備計の初期比率

            # 衛生設備の各内訳の初期比率（衛生設備計に対する比率）
            SANITARY_DETAIL_RATIOS = {
                19: 8.0,  # 衛生器具設備
                20: 8.0,  # 給水機器設備
                21: 8.0,  # 給水配管設備
                22: 11.0,  # 給湯機器設備
                23: 8.0,  # 給湯配管設備
                24: 19.0,  # 排水設備
                25: 6.0,  # ガス設備
                26: 5.0,  # 厨房設備
                27: 0.0,  # 浄化槽設備（初期値）
                28: 25.0,  # 消火設備
                29: 2.0,  # 雑排水設備
            }

            if septic_cost > 0:
                # 浄化槽がある場合
                new_cost = base_cost + septic_cost
                ratio_factor = new_cost / base_cost

                print(f"変更後の再調達価格: {new_cost:,.0f}円")
                print(f"比率係数: {ratio_factor:.6f}")

                # ========== 各工事区分の比率計算 ==========

                # 建築工事の比率計算
                new_arch_ratio = DEFAULT_ARCH_RATIO / ratio_factor
                self.table_vars[9][1].set(f"{new_arch_ratio:.2f}%")
                print(f"建築工事計の比率: {DEFAULT_ARCH_RATIO:.2f}% → {new_arch_ratio:.2f}%")

                # 空調設備工事の比率計算
                new_hvac_ratio = DEFAULT_HVAC_RATIO / ratio_factor
                self.table_vars[18][1].set(f"{new_hvac_ratio:.2f}%")
                print(f"空調設備工事計の比率: {DEFAULT_HVAC_RATIO:.2f}% → {new_hvac_ratio:.2f}%")

                # 電気設備工事の比率計算
                new_elec_ratio = DEFAULT_ELEC_RATIO / ratio_factor
                self.table_vars[38][1].set(f"{new_elec_ratio:.2f}%")
                print(f"電気設備計の比率: {DEFAULT_ELEC_RATIO:.2f}% → {new_elec_ratio:.2f}%")

                # 昇降機設備工事の比率計算
                new_elevator_ratio = DEFAULT_ELEVATOR_RATIO / ratio_factor
                self.table_vars[41][1].set(f"{new_elevator_ratio:.2f}%")
                print(f"昇降機設備計の比率: {DEFAULT_ELEVATOR_RATIO:.2f}% → {new_elevator_ratio:.2f}%")

                # ========== 衛生設備の計算 ==========
                print(f"\n--- 衛生設備の詳細 ---")

                # 衛生設備の各内訳の金額を計算（基本再調達価格ベース）
                old_sanitary_total_amount = base_cost * (DEFAULT_SANITARY_RATIO / 100)
                print(f"変更前の衛生設備計金額: {old_sanitary_total_amount:,.0f}円")

                # 変更後の衛生設備計金額 = 変更前の金額 + 浄化槽工事金額
                new_sanitary_total_amount = old_sanitary_total_amount + septic_cost
                print(f"変更後の衛生設備計金額: {new_sanitary_total_amount:,.0f}円")

                # 変更後の衛生設備計の比率
                new_sanitary_ratio = (new_sanitary_total_amount / new_cost) * 100
                self.table_vars[30][1].set(f"{new_sanitary_ratio:.2f}%")
                print(f"衛生設備計の比率: {DEFAULT_SANITARY_RATIO:.2f}% → {new_sanitary_ratio:.2f}%")

                # 各内訳項目の金額と比率を計算
                print(f"\n衛生設備の内訳:")
                for row, detail_ratio in SANITARY_DETAIL_RATIOS.items():
                    component_name = self.table_vars[row][0].get()

                    if row == 27:  # 浄化槽設備
                        # 浄化槽は追加される金額
                        component_amount = septic_cost
                        # 浄化槽の新築工事費全体に対する比率
                        component_ratio_to_total = (component_amount / new_cost) * 100
                        # 浄化槽の衛生設備計に対する比率
                        component_ratio_to_sanitary = (component_amount / new_sanitary_total_amount) * 100

                        self.table_vars[row][1].set(f"{component_ratio_to_sanitary:.2f}%")
                        print(f"  {component_name}: {component_amount:,.0f}円 "
                              f"(衛生設備計の{component_ratio_to_sanitary:.2f}%, "
                              f"新築工事費の{component_ratio_to_total:.2f}%)")
                    else:
                        # 他の衛生設備項目：金額は変わらないが比率が変わる
                        # 金額 = 変更前の衛生設備計金額 × 初期比率
                        component_amount = old_sanitary_total_amount * (detail_ratio / 100)
                        # 衛生設備計に対する新しい比率
                        component_ratio_to_sanitary = (component_amount / new_sanitary_total_amount) * 100
                        # 新築工事費全体に対する比率
                        component_ratio_to_total = (component_amount / new_cost) * 100

                        self.table_vars[row][1].set(f"{component_ratio_to_sanitary:.2f}%")
                        print(f"  {component_name}: {component_amount:,.0f}円 "
                              f"(衛生設備計の{component_ratio_to_sanitary:.2f}%, "
                              f"新築工事費の{component_ratio_to_total:.2f}%)")

            else:
                # 浄化槽がない場合は初期値に戻す
                print(f"浄化槽なし：比率を初期値に戻します")

                new_arch_ratio = DEFAULT_ARCH_RATIO
                new_hvac_ratio = DEFAULT_HVAC_RATIO
                new_sanitary_ratio = DEFAULT_SANITARY_RATIO
                new_elec_ratio = DEFAULT_ELEC_RATIO
                new_elevator_ratio = DEFAULT_ELEVATOR_RATIO

                # 比率を表に反映
                self.table_vars[9][1].set(f"{new_arch_ratio:.2f}%")
                self.table_vars[18][1].set(f"{new_hvac_ratio:.2f}%")
                self.table_vars[30][1].set(f"{new_sanitary_ratio:.2f}%")
                self.table_vars[38][1].set(f"{new_elec_ratio:.2f}%")
                self.table_vars[41][1].set(f"{new_elevator_ratio:.2f}%")

                # 衛生設備の各内訳も初期値に戻す
                for row, detail_ratio in SANITARY_DETAIL_RATIOS.items():
                    self.table_vars[row][1].set(f"{detail_ratio:.2f}%")

                # 浄化槽設備（行27）の金額欄（2列目）を"0"に設定
                self.table_vars[27][2].set("0")

            # 直接工事計の比率を計算
            total_ratio = new_arch_ratio + new_hvac_ratio + new_sanitary_ratio + new_elec_ratio + new_elevator_ratio
            self.table_vars[42][1].set(f"{total_ratio:.2f}%")
            print(f"\n直接工事計の比率: {total_ratio:.2f}%")

            print("=== 比率更新完了 ===\n")

            # LCC修繕計画表を再計算
            if hasattr(self, 'recalculate_table'):
                self.recalculate_table()

        except Exception as e:
            print(f"LCC比率更新エラー: {e}")
            import traceback
            traceback.print_exc()
    # コストサマリのタブに関する項目終了

    # 温泉設備のタブに関する項目開始
    def create_onsen_tab(self):
        """温泉設備タブの作成"""
        canvas = tk.Canvas(self.onsen_frame)
        scrollbar = ttk.Scrollbar(self.onsen_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(main_frame, text="温泉設備詳細設定", font= title_font).grid(row=0, column=0, columnspan=3,
                                                                                pady=(0, 20), sticky=tk.W)
        row = 1

        # 小計を追加
        subtotal_frame = ttk.Frame(main_frame)
        subtotal_frame.grid(row=row, column=0, sticky=tk.W, pady=8)
        ttk.Label(subtotal_frame, text="小計:", font=item_font).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Entry(subtotal_frame, textvariable=self.onsen_subtotal, width=10, state='readonly',
                  font=item_font, justify='right').pack(side=tk.LEFT)
        ttk.Label(subtotal_frame, text=" 円", font=item_font).pack(side=tk.LEFT, padx=(5, 0))
        row += 1

        ttk.Label(main_frame, text="浴室設定", font=sub_title_font).grid(row=row, column=0, columnspan=3,
                                                                                pady=(10, 5), sticky=tk.W)
        row += 1

        men_women_frame = ttk.Frame(main_frame)
        men_women_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        men_women_frame.grid_columnconfigure(0, weight=1)
        men_women_frame.grid_columnconfigure(1, weight=1)

        men_frame = ttk.LabelFrame(men_women_frame, text="男性用", padding="15")
        men_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10), pady=10)

        women_frame = ttk.LabelFrame(men_women_frame, text="女性用", padding="15")
        women_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0), pady=10)

        self.create_onsen_section(men_frame, "men")
        self.create_onsen_section(women_frame, "women")

        row += 1

        ttk.Label(main_frame, text="浴室関係設定 (計算に使用する変数)", font=sub_title_font).grid(row=row,
                                                                                         column=0,
                                                                                         columnspan=3,
                                                                                         pady=(20, 5),
                                                                                         sticky=tk.W)
        row += 1

        settings_input_frame = ttk.Frame(main_frame)
        settings_input_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E))
        #settings_input_frame.grid_columnconfigure(0, weight=1)
        #settings_input_frame.grid_columnconfigure(1, weight=1)

        left_col_frame = ttk.Frame(settings_input_frame)
        left_col_frame.grid(row=0, column=0, sticky=tk.N + tk.W, padx=(0, 10))

        right_col_frame = ttk.Frame(settings_input_frame)
        right_col_frame.grid(row=0, column=1, sticky=tk.N + tk.W)

        settings_data = [
            ("浴室営業時間:", self.bath_open_hours, "時間"),
            ("洗い場利用時間:", self.wash_time, "分"),
            ("浴場利用率:", self.bath_utilization_rate, "%"),
            ("浴槽利用時間:", self.tub_time, "分"),
            ("脱衣室利用時間:", self.changing_room_time, "分"),
            ("男洗面器利用時間:", self.men_face_wash_time, "分"),
            ("女洗面器利用時間:", self.women_face_wash_time, "分"),
            ("ピーク係数:", self.peak_factor, "2時間平均に対するピーク時間の係数"),
            ("利用者男性率:", self.men_ratio, "%"),
            ("利用者女性率:", self.women_ratio, "%"),
            ("洗い場使用面積:", self.wash_area_per_person, "㎡/人"),
            ("脱衣室使用面積:", self.changing_room_per_person, "㎡/人"),
            ("男脱衣室面積の割増:", self.men_changing_room_add, "%"),
            ("女脱衣室面積の割増:", self.women_changing_room_add, "%"),
            ("浴室内部通路割増:", self.bath_hallway_add, "%"),
            ("浴槽使用面積:", self.tub_area_per_person, "㎡/人"),
            ("浴室使用面積:", self.bath_area_per_person, "㎡/人"),
            ("浴室面積の割増:", self.bath_area_add, "%"),
            ("浴槽深さ:", self.tub_depth, "m"),
            ("洗面器利用率:", self.face_wash_utilization_rate, "%"),
            ("カラン利用率:", self.faucet_utilization_rate, "%"),
        ]

        half_point = math.ceil(len(settings_data) / 2)

        for i, (label_text, var, unit_text) in enumerate(settings_data):
            target_frame = left_col_frame if i < half_point else right_col_frame
            current_row = i if i < half_point else i - half_point

            ttk.Label(target_frame, text=label_text).grid(row=current_row, column=0, sticky=tk.W, pady=5)

            input_frame = ttk.Frame(target_frame)
            input_frame.grid(row=current_row, column=1, sticky=tk.W, padx=(5, 5))

            ttk.Entry(input_frame, textvariable=var, width=10).pack(side=tk.LEFT)
            ttk.Label(input_frame, text=unit_text).pack(side=tk.LEFT, padx=(5, 0))

        row += 1

        self.bath_open_hours.trace('w', self.update_onsen_calculations)#浴室営業時間
        self.wash_time.trace('w', self.update_onsen_calculations)#洗い場利用時間
        self.bath_utilization_rate.trace('w', self.update_onsen_calculations)#洗場利用率
        self.tub_time.trace('w', self.update_onsen_calculations)#浴槽利用時間
        self.changing_room_time.trace('w', self.update_onsen_calculations)#脱衣室利用時間
        self.men_face_wash_time.trace('w', self.update_onsen_calculations)#男洗面器利用時間
        self.women_face_wash_time.trace('w', self.update_onsen_calculations)#女洗面器利用時間
        self.peak_factor.trace('w', self.update_onsen_calculations)#ピーク係数
        self.men_ratio.trace('w', self.update_onsen_calculations)#利用者男性率
        self.women_ratio.trace('w', self.update_onsen_calculations)#利用者女性率
        self.wash_area_per_person.trace('w', self.update_onsen_calculations)#洗い場使用面積
        self.changing_room_per_person.trace('w', self.update_onsen_calculations)#脱衣室使用面積
        self.men_changing_room_add.trace('w', self.update_onsen_calculations)#男脱衣室面積の割増
        self.women_changing_room_add.trace('w', self.update_onsen_calculations)#女脱衣室面積の割増
        self.bath_hallway_add.trace('w', self.update_onsen_calculations)#浴室内部通路割増
        self.tub_area_per_person.trace('w', self.update_onsen_calculations)#浴槽使用面積
        self.bath_area_per_person.trace('w', self.update_onsen_calculations)#浴室使用面積
        self.bath_area_add.trace('w', self.update_onsen_calculations)#浴室面積の割増
        self.tub_depth.trace('w', self.update_onsen_calculations)#浴槽深さ
        self.face_wash_utilization_rate.trace('w', self.update_onsen_calculations)#洗面器利用率
        self.faucet_utilization_rate.trace('w', self.update_onsen_calculations)#カラン利用率
        self.men_indoor_bath_area.trace('w', self.update_onsen_bath_arch_costs)
        self.men_outdoor_bath_area.trace('w', self.update_onsen_bath_arch_costs)
        self.women_indoor_bath_area.trace('w', self.update_onsen_bath_arch_costs)
        self.women_outdoor_bath_area.trace('w', self.update_onsen_bath_arch_costs)

        self.men_bathroom_area.trace('w', self.update_onsen_bath_arch_costs)
        self.women_bathroom_area.trace('w', self.update_onsen_bath_arch_costs)

        # チェックボックスの変更時に小計を更新
        self.men_indoor_bath_check.trace('w', self.update_onsen_subtotal)
        self.men_outdoor_bath_check.trace('w', self.update_onsen_subtotal)
        self.men_bathroom_check.trace('w', self.update_onsen_subtotal)
        self.women_indoor_bath_check.trace('w', self.update_onsen_subtotal)
        self.women_outdoor_bath_check.trace('w', self.update_onsen_subtotal)
        self.women_bathroom_check.trace('w', self.update_onsen_subtotal)

        self.men_partition_check.trace('w', self.update_onsen_subtotal)
        self.men_wash_lighting_check.trace('w', self.update_onsen_subtotal)
        self.men_shower_faucet_check.trace('w', self.update_onsen_subtotal)
        self.women_partition_check.trace('w', self.update_onsen_subtotal)
        self.women_wash_lighting_check.trace('w', self.update_onsen_subtotal)
        self.women_shower_faucet_check.trace('w', self.update_onsen_subtotal)

        # 建築項目の金額が更新されたときに小計を更新
        self.men_indoor_bath_cost.trace('w', self.update_onsen_subtotal)
        self.men_outdoor_bath_cost.trace('w', self.update_onsen_subtotal)
        self.men_bathroom_cost.trace('w', self.update_onsen_subtotal)
        self.women_indoor_bath_cost.trace('w', self.update_onsen_subtotal)
        self.women_outdoor_bath_cost.trace('w', self.update_onsen_subtotal)
        self.women_bathroom_cost.trace('w', self.update_onsen_subtotal)

        # 洗い場設備の金額が更新されたときに小計を更新
        self.men_partition_cost.trace('w', self.update_onsen_subtotal)
        self.men_wash_lighting_cost.trace('w', self.update_onsen_subtotal)
        self.men_shower_faucet_cost.trace('w', self.update_onsen_subtotal)
        self.women_partition_cost.trace('w', self.update_onsen_subtotal)
        self.women_wash_lighting_cost.trace('w', self.update_onsen_subtotal)
        self.women_shower_faucet_cost.trace('w', self.update_onsen_subtotal)

    def create_onsen_section(self, parent_frame, gender):
        """温泉設備のセクション作成(男性用・女性用共通)"""
        parent_frame.grid_columnconfigure(0, weight=1)
        parent_frame.grid_columnconfigure(1, weight=1)

        # column 0 (ラベルの列) の重みを 0 に設定する→ ラベルが必要とする最小限の幅だけが確保され、余分なスペースは引き伸ばされない
        parent_frame.grid_columnconfigure(0, weight=0)

        row = 0

        ttk.Label(parent_frame, text="温泉タイプ", font=item_font).grid(row=row, column=0, columnspan=2,
                                                                                    pady=(5, 5), sticky=tk.W)
        row += 1

        # ========== ①内湯浴槽 ==========
        ttk.Label(parent_frame, text="内湯浴槽", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky=tk.W,
                                                                                  pady=5)
        row += 1

        # チェックボックスと面積入力
        indoor_check_var = tk.BooleanVar(value=False)
        setattr(self, f"{gender}_indoor_bath_check", indoor_check_var)

        ttk.Checkbutton(parent_frame, text="面積 (㎡):", variable=indoor_check_var).grid(
            row=row, column=0, sticky=tk.W, padx=(0, 10), pady=3)

        indoor_area_input_frame = ttk.Frame(parent_frame)
        indoor_area_input_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=3)

        indoor_area_var = tk.StringVar(value="0")
        ttk.Entry(indoor_area_input_frame, textvariable=indoor_area_var, width=10, justify=tk.RIGHT).pack(side=tk.LEFT)
        setattr(self, f"{gender}_indoor_bath_area", indoor_area_var)

        ttk.Label(indoor_area_input_frame, text="㎡").pack(side=tk.LEFT, padx=(5, 10))

        # 金額表示
        indoor_cost_var = tk.StringVar(value="0")
        ttk.Entry(indoor_area_input_frame, textvariable=indoor_cost_var, width=12, state='readonly',
                  justify=tk.RIGHT).pack(side=tk.LEFT, padx=(5, 5))
        ttk.Label(indoor_area_input_frame, text="円").pack(side=tk.LEFT)
        setattr(self, f"{gender}_indoor_bath_cost", indoor_cost_var)

        # 内湯容量表示(従来の推奨値)
        ttk.Label(indoor_area_input_frame, text="浴槽容量:", font=("Arial", 9)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(indoor_area_input_frame,
                  textvariable=getattr(self, f"{gender}_recommended_indoor_capacity"),
                  font=("Arial", 9, "bold"), foreground="blue").pack(side=tk.LEFT)#推奨値表示用の変数設定　から参照している
        ttk.Label(indoor_area_input_frame,
                  textvariable=getattr(self, f"{gender}_recommended_indoor_tub_area"),
                  font=("Arial", 9, "bold"), foreground="blue").pack(side=tk.LEFT)  # 内湯面積　推奨値表示用の変数設定　から参照している
        row += 1

        # ========== ②露天風呂 ==========
        ttk.Label(parent_frame, text="露天風呂", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky=tk.W,pady=5)
        row += 1

        # チェックボックスと面積入力
        outdoor_check_var = tk.BooleanVar(value=False)
        setattr(self, f"{gender}_outdoor_bath_check", outdoor_check_var)

        ttk.Checkbutton(parent_frame, text="面積 (㎡):", variable=outdoor_check_var).grid(
            row=row, column=0, sticky=tk.W, padx=(0, 10), pady=3)

        outdoor_area_input_frame = ttk.Frame(parent_frame)
        outdoor_area_input_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=3)

        outdoor_area_var = tk.StringVar(value="0")
        ttk.Entry(outdoor_area_input_frame, textvariable=outdoor_area_var, width=10, justify=tk.RIGHT).pack(side=tk.LEFT)
        setattr(self, f"{gender}_outdoor_bath_area", outdoor_area_var)

        ttk.Label(outdoor_area_input_frame, text="㎡").pack(side=tk.LEFT, padx=(5, 10))

        # 金額表示
        outdoor_cost_var = tk.StringVar(value="0")
        ttk.Entry(outdoor_area_input_frame, textvariable=outdoor_cost_var, width=12, state='readonly',
                  justify=tk.RIGHT).pack(side=tk.LEFT, padx=(5, 5))
        ttk.Label(outdoor_area_input_frame, text="円").pack(side=tk.LEFT)
        setattr(self, f"{gender}_outdoor_bath_cost", outdoor_cost_var)

        # 露天風呂容量表示(従来の推奨値)
        ttk.Label(outdoor_area_input_frame, text="浴槽容量:", font=("Arial", 9)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(outdoor_area_input_frame,
                  textvariable=getattr(self, f"{gender}_recommended_outdoor_capacity"),
                  font=("Arial", 9, "bold"), foreground="blue").pack(side=tk.LEFT)#推奨値表示用の変数設定　から参照している
        ttk.Label(outdoor_area_input_frame,
                  textvariable=getattr(self, f"{gender}_recommended_outdoor_tub_area"),
                  font=("Arial", 9, "bold"), foreground="blue").pack(side=tk.LEFT)  # 外湯面積　推奨値表示用の変数設定　から参照している
        row += 1

        # ========== ③浴室 ==========
        ttk.Label(parent_frame, text="浴室", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky=tk.W, pady=5)
        row += 1

        # チェックボックスと面積入力
        bathroom_check_var = tk.BooleanVar(value=False)
        setattr(self, f"{gender}_bathroom_check", bathroom_check_var)

        ttk.Checkbutton(parent_frame, text="面積 (㎡):", variable=bathroom_check_var).grid(
            row=row, column=0, sticky=tk.W, padx=(0, 10), pady=3)

        bathroom_area_input_frame = ttk.Frame(parent_frame)
        bathroom_area_input_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=3)

        # 面積入力(従来のself.{gender}_bathroom_areaを流用)
        bathroom_area_var = tk.StringVar(value="0")
        ttk.Entry(bathroom_area_input_frame, textvariable=bathroom_area_var, width=10, justify=tk.RIGHT).pack(
            side=tk.LEFT)
        setattr(self, f"{gender}_bathroom_area", bathroom_area_var)

        ttk.Label(bathroom_area_input_frame, text="㎡").pack(side=tk.LEFT, padx=(5, 10))

        # 金額表示
        bathroom_cost_var = tk.StringVar(value="0")
        ttk.Entry(bathroom_area_input_frame, textvariable=bathroom_cost_var, width=12, state='readonly',
                  justify=tk.RIGHT).pack(side=tk.LEFT, padx=(5, 5))
        ttk.Label(bathroom_area_input_frame, text="円").pack(side=tk.LEFT)
        setattr(self, f"{gender}_bathroom_cost", bathroom_cost_var)

        #row += 1

        # 推奨面積表示(従来の推奨値)
        #bathroom_recommended_label_frame = ttk.Frame(parent_frame)
        #bathroom_recommended_label_frame.grid(row=row, column=1, sticky=tk.W, pady=3)

        ttk.Label(bathroom_area_input_frame, text="推奨面積:", font=("Arial", 9)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(bathroom_area_input_frame,
                  textvariable=getattr(self, f"{gender}_recommended_bathroom_area"),
                  font=("Arial", 9, "bold"), foreground="blue").pack(side=tk.LEFT)

        row += 1

        # 建築項目の計算過程表示フレーム
        process_frame = ttk.LabelFrame(parent_frame, text="建築項目の計算過程", padding="10")
        process_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(15, 5))

        # 計算過程表示用の変数を作成
        arch_process_var = tk.StringVar(value="")
        setattr(self, f"{gender}_bath_arch_process", arch_process_var)

        process_label = ttk.Label(process_frame, textvariable=arch_process_var,
                                  font=("Arial", 9), justify=tk.LEFT, foreground="black")
        process_label.pack(fill=tk.BOTH, expand=True)

        row += 1

        #####################浴槽関連の計算過程を追加表示（ここを単独で消してもエラーは出ない）#############################
        # 推奨値の計算過程表示フレーム
        process_frame = ttk.LabelFrame(parent_frame, text="推奨値の計算式", padding="10")
        process_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(15, 5))

        process_label = ttk.Label(process_frame, textvariable=getattr(self, f"{gender}_calculation_process"),
                                  font=("Arial", 9), justify=tk.LEFT)
        process_label.pack(fill=tk.BOTH, expand=True)

        row += 1
        #####################浴槽関連の計算過程を追加表示（ここを単独で消してもエラーは出ない）#############################



        ttk.Label(parent_frame, text="洗い場", font=item_font).grid(row=row, column=0, columnspan=2,
                                                                                pady=(20, 5), sticky=tk.W)
        row += 1

        # 隔て板
        partition_frame = ttk.Frame(parent_frame)
        partition_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)

        partition_check_var = getattr(self, f"{gender}_partition_check")
        ttk.Checkbutton(partition_frame, text="隔て板", variable=partition_check_var, width=12).grid(row=0, column=0,
                                                                                                     sticky=tk.W)

        ttk.Label(partition_frame, text="個数:").grid(row=0, column=1, sticky=tk.E, padx=(5, 5))
        partition_count_var = getattr(self, f"{gender}_partition_count")
        ttk.Entry(partition_frame, textvariable=partition_count_var, width=8).grid(row=0, column=2, sticky=tk.W)

        partition_recommended_var = getattr(self, f"{gender}_partition_recommended")
        ttk.Label(partition_frame, textvariable=partition_recommended_var, font=("Arial", 9, "bold"), foreground="blue",
                  width=12, anchor=tk.W).grid(row=0, column=3, sticky=tk.W, padx=(5, 5))

        ttk.Label(partition_frame, text="金額:").grid(row=0, column=4, sticky=tk.E, padx=(10, 5))
        partition_cost_var = getattr(self, f"{gender}_partition_cost")
        ttk.Entry(partition_frame, textvariable=partition_cost_var, width=12, state='readonly', justify=tk.RIGHT).grid(
            row=0, column=5, sticky=tk.W)
        ttk.Label(partition_frame, text="円").grid(row=0, column=6, sticky=tk.W, padx=(5, 0))

        row += 1

        # 照明
        lighting_frame = ttk.Frame(parent_frame)
        lighting_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)

        lighting_check_var = getattr(self, f"{gender}_wash_lighting_check")
        ttk.Checkbutton(lighting_frame, text="照明", variable=lighting_check_var, width=12).grid(row=0, column=0,
                                                                                                 sticky=tk.W)

        ttk.Label(lighting_frame, text="個数:").grid(row=0, column=1, sticky=tk.E, padx=(5, 5))
        lighting_count_var = getattr(self, f"{gender}_wash_lighting_count")
        ttk.Entry(lighting_frame, textvariable=lighting_count_var, width=8).grid(row=0, column=2, sticky=tk.W)

        lighting_recommended_var = getattr(self, f"{gender}_wash_lighting_recommended")
        ttk.Label(lighting_frame, textvariable=lighting_recommended_var, font=("Arial", 9, "bold"), foreground="blue",
                  width=12, anchor=tk.W).grid(row=0, column=3, sticky=tk.W, padx=(5, 5))

        ttk.Label(lighting_frame, text="金額:").grid(row=0, column=4, sticky=tk.E, padx=(10, 5))
        lighting_cost_var = getattr(self, f"{gender}_wash_lighting_cost")
        ttk.Entry(lighting_frame, textvariable=lighting_cost_var, width=12, state='readonly', justify=tk.RIGHT).grid(
            row=0, column=5, sticky=tk.W)
        ttk.Label(lighting_frame, text="円").grid(row=0, column=6, sticky=tk.W, padx=(5, 0))

        row += 1

        # シャワー水栓
        shower_frame = ttk.Frame(parent_frame)
        shower_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)

        shower_check_var = getattr(self, f"{gender}_shower_faucet_check")
        ttk.Checkbutton(shower_frame, text="シャワー水栓", variable=shower_check_var, width=12).grid(row=0, column=0,
                                                                                                     sticky=tk.W)

        ttk.Label(shower_frame, text="個数:").grid(row=0, column=1, sticky=tk.E, padx=(5, 5))
        shower_count_var = getattr(self, f"{gender}_shower_faucet_count")
        ttk.Entry(shower_frame, textvariable=shower_count_var, width=8).grid(row=0, column=2, sticky=tk.W)

        shower_recommended_var = getattr(self, f"{gender}_shower_faucet_recommended")
        ttk.Label(shower_frame, textvariable=shower_recommended_var, font=("Arial", 9, "bold"), foreground="blue",
                  width=12, anchor=tk.W).grid(row=0, column=3, sticky=tk.W, padx=(5, 5))

        ttk.Label(shower_frame, text="金額:").grid(row=0, column=4, sticky=tk.E, padx=(10, 5))
        shower_cost_var = getattr(self, f"{gender}_shower_faucet_cost")
        ttk.Entry(shower_frame, textvariable=shower_cost_var, width=12, state='readonly', justify=tk.RIGHT).grid(row=0,
                                                                                                                 column=5,
                                                                                                                 sticky=tk.W)
        ttk.Label(shower_frame, text="円").grid(row=0, column=6, sticky=tk.W, padx=(5, 0))

        row += 1

        # 洗い場設備の計算過程表示フレーム
        wash_process_frame = ttk.LabelFrame(parent_frame, text="洗い場設備の計算過程", padding="10")
        wash_process_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 5))

        wash_process_var = getattr(self, f"{gender}_wash_equipment_process")
        wash_process_label = ttk.Label(wash_process_frame, textvariable=wash_process_var, font=("Courier", 9),
                                       justify=tk.LEFT, foreground="darkgreen")
        wash_process_label.pack(fill=tk.BOTH, expand=True)

        row += 1

        ttk.Label(parent_frame, text="特殊設備", font=item_font).grid(row=row, column=0, columnspan=2,
                                                                                  pady=(20, 5), sticky=tk.W)
        row += 1

        # サウナ
        sauna_frame = ttk.Frame(parent_frame)
        sauna_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)

        sauna_var = getattr(self, f"{gender}_sauna")
        ttk.Checkbutton(sauna_frame, text="サウナ", variable=sauna_var, width=12).grid(row=0, column=0, sticky=tk.W)

        ttk.Label(sauna_frame, text="金額:").grid(row=0, column=1, sticky=tk.E, padx=(10, 5))
        sauna_cost_var = getattr(self, f"{gender}_sauna_cost")
        ttk.Entry(sauna_frame, textvariable=sauna_cost_var, width=12, state='readonly', justify=tk.RIGHT).grid(
            row=0, column=2, sticky=tk.W)
        ttk.Label(sauna_frame, text="円").grid(row=0, column=3, sticky=tk.W, padx=(5, 0))

        row += 1

        # ジャグジー
        jacuzzi_frame = ttk.Frame(parent_frame)
        jacuzzi_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)

        jacuzzi_var = getattr(self, f"{gender}_jacuzzi")
        ttk.Checkbutton(jacuzzi_frame, text="ジャグジー", variable=jacuzzi_var, width=12).grid(row=0, column=0,
                                                                                               sticky=tk.W)

        ttk.Label(jacuzzi_frame, text="金額:").grid(row=0, column=1, sticky=tk.E, padx=(10, 5))
        jacuzzi_cost_var = getattr(self, f"{gender}_jacuzzi_cost")
        ttk.Entry(jacuzzi_frame, textvariable=jacuzzi_cost_var, width=12, state='readonly', justify=tk.RIGHT).grid(
            row=0, column=2, sticky=tk.W)
        ttk.Label(jacuzzi_frame, text="円").grid(row=0, column=3, sticky=tk.W, padx=(5, 0))

        row += 1

        # 水風呂
        water_bath_frame = ttk.Frame(parent_frame)
        water_bath_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)

        water_bath_var = getattr(self, f"{gender}_water_bath")
        ttk.Checkbutton(water_bath_frame, text="水風呂", variable=water_bath_var, width=12).grid(row=0, column=0,
                                                                                                 sticky=tk.W)

        ttk.Label(water_bath_frame, text="金額:").grid(row=0, column=1, sticky=tk.E, padx=(10, 5))
        water_bath_cost_var = getattr(self, f"{gender}_water_bath_cost")
        ttk.Entry(water_bath_frame, textvariable=water_bath_cost_var, width=12, state='readonly',
                  justify=tk.RIGHT).grid(
            row=0, column=2, sticky=tk.W)
        ttk.Label(water_bath_frame, text="円").grid(row=0, column=3, sticky=tk.W, padx=(5, 0))

        row += 1

    def update_wash_equipment_recommended(self, *args):
        """洗い場設備の推奨値を計算"""
        try:
            total_guests = self.get_total_guests()
            bath_open_hours = self.bath_open_hours.get()
            peak_factor = self.peak_factor.get()
            men_ratio = self.men_ratio.get()
            women_ratio = self.women_ratio.get()
            tub_time = self.tub_time.get()

            if bath_open_hours > 0:
                # 男性用の推奨値計算
                men_recommended = math.ceil((men_ratio / 100) * total_guests / bath_open_hours * peak_factor/ bath_open_hours*tub_time/60)
                self.men_partition_recommended.set(f"推奨: {men_recommended}個")
                self.men_wash_lighting_recommended.set(f"推奨: {men_recommended}個")
                self.men_shower_faucet_recommended.set(f"推奨: {men_recommended}個")

                # 女性用の推奨値計算
                women_recommended = math.ceil((women_ratio / 100) * total_guests / bath_open_hours * peak_factor/ bath_open_hours*tub_time/60)
                self.women_partition_recommended.set(f"推奨: {women_recommended}個")
                self.women_wash_lighting_recommended.set(f"推奨: {women_recommended}個")
                self.women_shower_faucet_recommended.set(f"推奨: {women_recommended}個")
            else:
                self.men_partition_recommended.set("")
                self.men_wash_lighting_recommended.set("")
                self.men_shower_faucet_recommended.set("")
                self.women_partition_recommended.set("")
                self.women_wash_lighting_recommended.set("")
                self.women_shower_faucet_recommended.set("")

        except (ValueError, ZeroDivisionError):
            self.men_partition_recommended.set("")
            self.men_wash_lighting_recommended.set("")
            self.men_shower_faucet_recommended.set("")
            self.women_partition_recommended.set("")
            self.women_wash_lighting_recommended.set("")
            self.women_shower_faucet_recommended.set("")

    def update_wash_equipment_costs(self, *args):
        """洗い場設備の金額を計算"""
        # 男性用の計算過程テキスト
        men_process_text = "【男性用 洗い場設備の計算】\n"

        # 男性用 隔て板
        try:
            count = float(self.men_partition_count.get()) if self.men_partition_count.get() else 0
            if count > 0:
                # 設備設定から隔て板の値を取得
                data = self.get_equipment_data_by_abbrev("隔て板")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.men_partition_cost.set(f"{int(cost):,}")
                    men_process_text += f"\n[隔て板]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
                else:
                    self.men_partition_cost.set("0")
            else:
                self.men_partition_cost.set("0")
        except (ValueError, TypeError):
            self.men_partition_cost.set("0")

        # 男性用 照明
        try:
            count = float(self.men_wash_lighting_count.get()) if self.men_wash_lighting_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("洗い場照明")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.men_wash_lighting_cost.set(f"{int(cost):,}")
                    men_process_text += f"\n[照明]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
                else:
                    self.men_wash_lighting_cost.set("0")
            else:
                self.men_wash_lighting_cost.set("0")
        except (ValueError, TypeError):
            self.men_wash_lighting_cost.set("0")

        # 男性用 シャワー水栓
        try:
            count = float(self.men_shower_faucet_count.get()) if self.men_shower_faucet_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("シャワー水栓")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.men_shower_faucet_cost.set(f"{int(cost):,}")
                    men_process_text += f"\n[シャワー水栓]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
                else:
                    self.men_shower_faucet_cost.set("0")
            else:
                self.men_shower_faucet_cost.set("0")
        except (ValueError, TypeError):
            self.men_shower_faucet_cost.set("0")

        self.men_wash_equipment_process.set(men_process_text)

        # 女性用の計算過程テキスト
        women_process_text = "【女性用 洗い場設備の計算】\n"

        # 女性用 隔て板
        try:
            count = float(self.women_partition_count.get()) if self.women_partition_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("隔て板")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.women_partition_cost.set(f"{int(cost):,}")
                    women_process_text += f"\n[隔て板]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
                else:
                    self.women_partition_cost.set("0")
            else:
                self.women_partition_cost.set("0")
        except (ValueError, TypeError):
            self.women_partition_cost.set("0")

        # 女性用 照明
        try:
            count = float(self.women_wash_lighting_count.get()) if self.women_wash_lighting_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("洗い場照明")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.women_wash_lighting_cost.set(f"{int(cost):,}")
                    women_process_text += f"\n[照明]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
                else:
                    self.women_wash_lighting_cost.set("0")
            else:
                self.women_wash_lighting_cost.set("0")
        except (ValueError, TypeError):
            self.women_wash_lighting_cost.set("0")

        # 女性用 シャワー水栓
        try:
            count = float(self.women_shower_faucet_count.get()) if self.women_shower_faucet_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("シャワー水栓")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.women_shower_faucet_cost.set(f"{int(cost):,}")
                    women_process_text += f"\n[シャワー水栓]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
                else:
                    self.women_shower_faucet_cost.set("0")
            else:
                self.women_shower_faucet_cost.set("0")
        except (ValueError, TypeError):
            self.women_shower_faucet_cost.set("0")

        self.women_wash_equipment_process.set(women_process_text)
        self.update_onsen_subtotal()

    def update_special_equipment_costs(self, *args):
        """特殊設備の金額を計算"""
        try:
            total_guests = self.get_total_guests()
            men_ratio = self.men_ratio.get()
            women_ratio = self.women_ratio.get()

            # 男性用の客数
            men_guests = total_guests * (men_ratio / 100)

            # 男性用
            # サウナ
            if self.men_sauna.get():
                cost = men_guests * 200 + 1500000
                self.men_sauna_cost.set(f"{int(cost):,}")
            else:
                self.men_sauna_cost.set("0")

            # ジャグジー
            if self.men_jacuzzi.get():
                cost = men_guests * 500 + 3000000
                self.men_jacuzzi_cost.set(f"{int(cost):,}")
            else:
                self.men_jacuzzi_cost.set("0")

            # 水風呂
            if self.men_water_bath.get():
                cost = men_guests * 500 + 1000000
                self.men_water_bath_cost.set(f"{int(cost):,}")
            else:
                self.men_water_bath_cost.set("0")

            # 女性用の客数
            women_guests = total_guests * (women_ratio / 100)

            # 女性用
            # サウナ
            if self.women_sauna.get():
                cost = women_guests * 200 + 1500000
                self.women_sauna_cost.set(f"{int(cost):,}")
            else:
                self.women_sauna_cost.set("0")

            # ジャグジー
            if self.women_jacuzzi.get():
                cost = women_guests * 500 + 3000000
                self.women_jacuzzi_cost.set(f"{int(cost):,}")
            else:
                self.women_jacuzzi_cost.set("0")

            # 水風呂
            if self.women_water_bath.get():
                cost = women_guests * 500 + 1000000
                self.women_water_bath_cost.set(f"{int(cost):,}")
            else:
                self.women_water_bath_cost.set("0")

        except (ValueError, AttributeError):
            # エラー時は0にリセット
            self.men_sauna_cost.set("0")
            self.men_jacuzzi_cost.set("0")
            self.men_water_bath_cost.set("0")
            self.women_sauna_cost.set("0")
            self.women_jacuzzi_cost.set("0")
            self.women_water_bath_cost.set("0")

    def update_onsen_calculations(self, *args):
        """温泉設備の計算を更新する"""
        try:
            total_guests = self.get_total_guests()

            peak_factor = self.peak_factor.get()
            tub_area_per_person = self.tub_area_per_person.get()
            bath_area_per_person = self.bath_area_per_person.get()
            tub_depth = self.tub_depth.get()
            bath_area_add = self.bath_area_add.get()
            bath_open_hours = self.bath_open_hours.get()
            tub_time=self.tub_time.get()
            wash_time=self.wash_time.get()

            men_ratio = self.men_ratio.get()
            wash_area_per_person= self.wash_area_per_person.get()#洗い場の面積を使ってない


            men_process_text = (

                f"内湯容量(m³) = (人数) × (比率) × (ピーク係数) ÷(浴室営業時間)× (浴室利用時間÷60) × (一人当浴槽面積) ×(深さ)\n"
                f"　　　　　　 = {total_guests}  × ({men_ratio}/100) × {peak_factor}  ÷ {bath_open_hours} × {tub_time} ÷60 × {tub_area_per_person}  × {tub_depth}  "
                f"= {total_guests * (men_ratio / 100) * peak_factor / bath_open_hours*tub_time/60* tub_area_per_person * tub_depth :.2f} m³\n"
                
                f"露天風呂容量(m³) = ({total_guests * (men_ratio / 100)* peak_factor / bath_open_hours*tub_time/60 * tub_area_per_person * tub_depth :.2f} / 2) "
                f"={total_guests * (men_ratio / 100)* peak_factor / bath_open_hours*tub_time/60 * tub_area_per_person * tub_depth/2 :.2f}m³　(内湯の50％として)\n"
                
                f"浴室面積(㎡) = 〔(人数) × (比率) × (ピーク係数)〕÷ (浴室営業時間) × (浴室利用時間÷60) × (一人当面積)× (浴室面積の割増/100)\n"
                f"　　　　　　 = {total_guests}  × ({men_ratio}/100) × {peak_factor} / {bath_open_hours} × {tub_time} ÷60÷{bath_area_per_person}× ({bath_area_add}/100)"
                f"={total_guests * (men_ratio / 100) *  peak_factor /bath_open_hours*(tub_time+wash_time)/60*bath_area_per_person* (bath_area_add / 100):.2f} ㎡\n"


            )

            self.men_calculation_process.set(men_process_text)

            men_total_tub_volume_m3 = total_guests * (men_ratio / 100) * peak_factor / bath_open_hours*tub_time/60* tub_area_per_person * tub_depth
            men_recommended_indoor_capacity_lit = men_total_tub_volume_m3  * 1000# 推奨値表示用の変数設定　に参照している
            men_recommended_indoor_tub_area = men_total_tub_volume_m3 /tub_depth  #(内湯面積) 推奨値表示用の変数設定　に参照している
            men_recommended_outdoor_capacity_lit = (men_total_tub_volume_m3 / 2) * 1000
            men_recommended_outdoor_tub_area = men_total_tub_volume_m3 / tub_depth /2 # (外湯面積) 推奨値表示用の変数設定　に参照している
            men_recommended_bath_area_m2 = total_guests * (men_ratio / 100) * peak_factor/ bath_open_hours*(tub_time+wash_time) / 60* bath_area_per_person  * (bath_area_add / 100)

            # 推奨値表示用の変数設定
            self.men_recommended_indoor_capacity.set(f"推奨値: {int(men_recommended_indoor_capacity_lit)} L")
            self.men_recommended_indoor_tub_area.set(f"／推奨値: {int(men_recommended_indoor_tub_area)} m2")
            self.men_recommended_outdoor_capacity.set(f"推奨値: {int(men_recommended_outdoor_capacity_lit)} L")
            self.men_recommended_outdoor_tub_area.set(f"／推奨値: {int(men_recommended_outdoor_tub_area)} m2")
            self.men_recommended_bathroom_area.set(f"推奨値: {int(men_recommended_bath_area_m2)} ㎡")

            women_ratio = self.women_ratio.get()

            women_process_text = (

                f"内湯容量(m³) = (人数) × (比率) × (ピーク係数) ÷(浴室営業時間)× (浴室利用時間÷60) × (一人当浴槽面積) ×(深さ)\n"
                f"　　　　　　 = {total_guests}  × ({women_ratio}/100) × {peak_factor}  ÷ {bath_open_hours} × {tub_time} ÷60 × {tub_area_per_person}  × {tub_depth}  "
                f"= {total_guests * (women_ratio / 100) * peak_factor / bath_open_hours * tub_time / 60 * tub_area_per_person * tub_depth :.2f} m³\n"

                f"露天風呂容量(m³) = ({total_guests * (women_ratio / 100) * peak_factor / bath_open_hours * tub_time / 60 * tub_area_per_person * tub_depth :.2f} / 2) "
                f"={total_guests * (women_ratio / 100) * peak_factor / bath_open_hours * tub_time / 60 * tub_area_per_person * tub_depth / 2 :.2f}m³　(内湯の50％として)\n"

                f"浴室面積(㎡) = 〔(人数) × (比率) × (ピーク係数)〕÷ (浴室営業時間) × (浴室利用時間÷60) × (一人当面積)× (浴室面積の割増/100)\n"
                f"　　　　　　 = {total_guests}  × ({women_ratio}/100) × {peak_factor} / {bath_open_hours} × {tub_time} ÷60÷{bath_area_per_person}× ({bath_area_add}/100)"
                f"={total_guests * (women_ratio / 100) * peak_factor / bath_open_hours * (tub_time + wash_time) / 60 * bath_area_per_person * (bath_area_add / 100):.2f} ㎡\n"

            )

            self.women_calculation_process.set(women_process_text)

            women_total_tub_volume_m3 = total_guests * (women_ratio / 100) * peak_factor / bath_open_hours*tub_time/60* tub_area_per_person * tub_depth
            women_recommended_indoor_capacity_lit = women_total_tub_volume_m3  * 1000
            women_recommended_indoor_tub_area = women_total_tub_volume_m3 / tub_depth  # (内湯面積) 推奨値表示用の変数設定　に参照している
            women_recommended_outdoor_capacity_lit = (women_total_tub_volume_m3 / 2) * 1000
            women_recommended_outdoor_tub_area = women_total_tub_volume_m3 / tub_depth / 2  # (外湯面積) 推奨値表示用の変数設定　に参照している
            women_recommended_bath_area_m2 = total_guests * (women_ratio / 100) * peak_factor / bath_open_hours * (tub_time + wash_time) / 60 * bath_area_per_person * (bath_area_add / 100)

            # 推奨値表示用の変数設定
            self.women_recommended_indoor_capacity.set(f"推奨値: {int(women_recommended_indoor_capacity_lit)} L")
            self.women_recommended_indoor_tub_area.set(f"／推奨値: {int(women_recommended_indoor_tub_area)} m2")
            self.women_recommended_outdoor_capacity.set(f"推奨値: {int(women_recommended_outdoor_capacity_lit)} L")
            self.women_recommended_outdoor_tub_area.set(f"／推奨値: {int(women_recommended_outdoor_tub_area)} m2")
            self.women_recommended_bathroom_area.set(f"推奨値: {int(women_recommended_bath_area_m2)} ㎡")


            self.update_wash_equipment_recommended()# 洗い場設備の推奨値を計算
            self.update_special_equipment_costs()  # 特殊設備の推奨値計算

        except (ValueError, ZeroDivisionError):
            self.men_recommended_indoor_capacity.set("推奨値: N/A")
            self.men_recommended_outdoor_capacity.set("推奨値: N/A")
            self.women_recommended_indoor_capacity.set("推奨値: N/A")
            self.women_recommended_outdoor_capacity.set("推奨値: N/A")
            self.men_recommended_bathroom_area.set("推奨値: N/A")
            self.women_recommended_bathroom_area.set("推奨値: N/A")
            self.men_calculation_process.set("計算に必要な値が不正です。")
            self.women_calculation_process.set("計算に必要な値が不正です。")

    def update_onsen_bath_arch_costs(self, *args):
        """温泉設備タブの建築項目の金額を自動計算"""

        # ========== 男性用 ==========
        men_process_text = "【男性用 建築項目の計算】\n"

        # ①内湯浴槽
        try:
            area = float(self.men_indoor_bath_area.get())
            if area > 0:
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("内風呂")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    self.men_indoor_bath_cost.set(f"{int(cost):,}")

                    men_process_text += (f"\n[内湯浴槽 - {grade_name}]\n"
                                         f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                                         f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                                         f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                                         f"       = {int(cost):,}円\n")
                else:
                    self.men_indoor_bath_cost.set("0")
                    men_process_text += "\n[内湯浴槽]単価情報が取得できません\n"
            else:
                self.men_indoor_bath_cost.set("0")
        except ValueError:
            self.men_indoor_bath_cost.set("0")

        # ②露天風呂
        try:
            area = float(self.men_outdoor_bath_area.get())
            if area > 0:
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("露天風呂")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    self.men_outdoor_bath_cost.set(f"{int(cost):,}")

                    men_process_text += (f"\n[露天風呂 - {grade_name}]\n"
                                         f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                                         f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                                         f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                                         f"       = {int(cost):,}円\n")
                else:
                    self.men_outdoor_bath_cost.set("0")
                    men_process_text += "\n[露天風呂]単価情報が取得できません\n"
            else:
                self.men_outdoor_bath_cost.set("0")
        except ValueError:
            self.men_outdoor_bath_cost.set("0")

        # ③浴室
        try:
            area = float(self.men_bathroom_area.get())
            if area > 0:
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("浴室")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    self.men_bathroom_cost.set(f"{int(cost):,}")

                    men_process_text += (f"\n[浴室 - {grade_name}]\n"
                                         f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                                         f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                                         f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                                         f"       = {int(cost):,}円\n")
                else:
                    self.men_bathroom_cost.set("0")
                    men_process_text += "\n[浴室]単価情報が取得できません\n"
            else:
                self.men_bathroom_cost.set("0")
        except ValueError:
            self.men_bathroom_cost.set("0")

        self.men_bath_arch_process.set(men_process_text)

        # ========== 女性用 ==========
        women_process_text = "【女性用 建築項目の計算】\n"

        # ①内湯浴槽
        try:
            area = float(self.women_indoor_bath_area.get())
            if area > 0:
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("内風呂")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    self.women_indoor_bath_cost.set(f"{int(cost):,}")

                    women_process_text += (f"\n[内湯浴槽 - {grade_name}]\n"
                                           f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                                           f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                                           f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                                           f"       = {int(cost):,}円\n")
                else:
                    self.women_indoor_bath_cost.set("0")
                    women_process_text += "\n[内湯浴槽]単価情報が取得できません\n"
            else:
                self.women_indoor_bath_cost.set("0")
        except ValueError:
            self.women_indoor_bath_cost.set("0")

        # ②露天風呂
        try:
            area = float(self.women_outdoor_bath_area.get())
            if area > 0:
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("露天風呂")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    self.women_outdoor_bath_cost.set(f"{int(cost):,}")

                    women_process_text += (f"\n[露天風呂 - {grade_name}]\n"
                                           f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                                           f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                                           f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                                           f"       = {int(cost):,}円\n")
                else:
                    self.women_outdoor_bath_cost.set("0")
                    women_process_text += "\n[露天風呂]単価情報が取得できません\n"
            else:
                self.women_outdoor_bath_cost.set("0")
        except ValueError:
            self.women_outdoor_bath_cost.set("0")

        # ③浴室
        try:
            area = float(self.women_bathroom_area.get())
            if area > 0:
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("浴室")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    self.women_bathroom_cost.set(f"{int(cost):,}")

                    women_process_text += (f"\n[浴室 - {grade_name}]\n"
                                           f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                                           f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                                           f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                                           f"       = {int(cost):,}円\n")
                else:
                    self.women_bathroom_cost.set("0")
                    women_process_text += "\n[浴室]単価情報が取得できません\n"
            else:
                self.women_bathroom_cost.set("0")
        except ValueError:
            self.women_bathroom_cost.set("0")

        self.women_bath_arch_process.set(women_process_text)

        # 小計を更新
        self.update_onsen_subtotal()

    def update_onsen_subtotal(self, *args):
        """温泉設備タブの小計を計算"""
        total = 0

        # 男性用の建築項目
        if self.men_indoor_bath_check.get():
            try:
                cost_str = self.men_indoor_bath_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        if self.men_outdoor_bath_check.get():
            try:
                cost_str = self.men_outdoor_bath_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        if self.men_bathroom_check.get():
            try:
                cost_str = self.men_bathroom_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        # 女性用の建築項目
        if self.women_indoor_bath_check.get():
            try:
                cost_str = self.women_indoor_bath_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        if self.women_outdoor_bath_check.get():
            try:
                cost_str = self.women_outdoor_bath_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        if self.women_bathroom_check.get():
            try:
                cost_str = self.women_bathroom_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        # 男性用の洗い場設備
        if self.men_partition_check.get():
            try:
                cost_str = self.men_partition_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        if self.men_wash_lighting_check.get():
            try:
                cost_str = self.men_wash_lighting_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        if self.men_shower_faucet_check.get():
            try:
                cost_str = self.men_shower_faucet_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        # 女性用の洗い場設備
        if self.women_partition_check.get():
            try:
                cost_str = self.women_partition_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        if self.women_wash_lighting_check.get():
            try:
                cost_str = self.women_wash_lighting_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        if self.women_shower_faucet_check.get():
            try:
                cost_str = self.women_shower_faucet_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        if self.women_shower_faucet_check.get():
            try:
                cost_str = self.women_shower_faucet_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

            # ========== 以下を追加 ==========
            # 男性用の特殊設備
        if self.men_sauna.get():
            try:
                cost_str = self.men_sauna_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        if self.men_jacuzzi.get():
            try:
                cost_str = self.men_jacuzzi_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        if self.men_water_bath.get():
            try:
                cost_str = self.men_water_bath_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

            # 女性用の特殊設備
        if self.women_sauna.get():
            try:
                cost_str = self.women_sauna_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        if self.women_jacuzzi.get():
            try:
                cost_str = self.women_jacuzzi_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        if self.women_water_bath.get():
            try:
                cost_str = self.women_water_bath_cost.get().replace(',', '')
                total += float(cost_str) if cost_str else 0
            except ValueError:
                pass

        # 小計を更新
        self.onsen_subtotal.set(f"{int(total):,}")
    # 温泉設備のタブに関する項目終了

    # レストランのタブに関する項目開始
    def create_restaurant_tab(self):
        """レストランタブの作成"""
        canvas = tk.Canvas(self.restaurant_frame)
        scrollbar = ttk.Scrollbar(self.restaurant_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(main_frame, text="レストラン詳細設定", font=title_font).grid(row=0, column=0, columnspan=3,
                                                                               pady=(0, 20), sticky=tk.W)
        ttk.Label(main_frame, text="面積を入力すると金額が表示されます。計上金額に含める場合は、項目に☑を入れてください。",
                  font=("Arial", 8)).grid(row=1, column=0, columnspan=3, sticky=tk.E, pady=(20, 0))
        row = 2

        # レストラン小計
        subtotal_frame = ttk.Frame(main_frame)
        subtotal_frame.grid(row=row, column=0, sticky=tk.W, pady=8)
        ttk.Label(subtotal_frame, text="小計:", font=item_font).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Entry(subtotal_frame, textvariable=self.restaurant_subtotal, width=10, state='readonly',
                  font=item_font, justify='right').pack(side=tk.LEFT)
        ttk.Label(subtotal_frame, text=" 円", font=item_font).pack(side=tk.LEFT, padx=(5, 0))
        row += 1

        # ①建築フレーム
        arch_frame = ttk.LabelFrame(main_frame, text="建築", padding="15")
        arch_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1

        self.add_lounge_item(arch_frame, 0, "レストラン", self.restaurant_arch_restaurant_check,
                              self.restaurant_arch_restaurant_area, "㎡", self.restaurant_arch_restaurant_cost)
        self.add_lounge_item(arch_frame, 1, "ライブキッチン", self.restaurant_arch_livekitchen_check,
                              self.restaurant_arch_livekitchen_area, "㎡", self.restaurant_arch_livekitchen_cost)
        self.add_lounge_item(arch_frame, 2, "厨房", self.restaurant_arch_kitchen_check,
                              self.restaurant_arch_kitchen_area, "㎡", self.restaurant_arch_kitchen_cost)

        # 建築項目の計算過程
        arch_process_frame = ttk.LabelFrame(main_frame, text="建築項目の計算過程", padding="10")
        arch_process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        arch_process_text_frame = ttk.Frame(arch_process_frame)
        arch_process_text_frame.pack(fill=tk.BOTH, expand=True)

        restaurant_label = ttk.Label(arch_process_text_frame, textvariable=self.restaurant_arch_restaurant_process,
                                     font=("Courier", 9), justify=tk.LEFT, foreground="blue")
        restaurant_label.pack(anchor=tk.W, pady=3)

        livekitchen_label = ttk.Label(arch_process_text_frame, textvariable=self.restaurant_arch_livekitchen_process,
                                      font=("Courier", 9), justify=tk.LEFT, foreground="blue")
        livekitchen_label.pack(anchor=tk.W, pady=3)

        kitchen_label = ttk.Label(arch_process_text_frame, textvariable=self.restaurant_arch_kitchen_process,
                                  font=("Courier", 9), justify=tk.LEFT, foreground="blue")
        kitchen_label.pack(anchor=tk.W, pady=3)

        # ③機械設備フレーム
        mech_frame = ttk.LabelFrame(main_frame, text="機械設備", padding="15")
        mech_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
        row += 1

        self.add_lounge_item(mech_frame, 0, "共用部エアコン", self.restaurant_mech_ac_check,
                              self.restaurant_mech_ac_count, "台", self.restaurant_mech_ac_cost,
                              self.restaurant_mech_ac_recommended)
        self.add_lounge_item(mech_frame, 1, "スプリンクラー", self.restaurant_mech_sprinkler_check,
                              self.restaurant_mech_sprinkler_count, "個", self.restaurant_mech_sprinkler_cost,
                              self.restaurant_mech_sprinkler_recommended)
        self.add_lounge_item(mech_frame, 2, "消火栓箱", self.restaurant_mech_fire_hose_check,
                              self.restaurant_mech_fire_hose_count, "台", self.restaurant_mech_fire_hose_cost,
                              self.restaurant_mech_fire_hose_recommended)
        self.add_lounge_item(mech_frame, 3, "フード", self.restaurant_mech_hood_check,
                              self.restaurant_mech_hood_count, "台", self.restaurant_mech_hood_cost,
                              self.restaurant_mech_hood_recommended)

        # 機械設備の計算過程
        mech_process_frame = ttk.LabelFrame(main_frame, text="機械設備の計算過程", padding="10")
        mech_process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        mech_process_label = ttk.Label(mech_process_frame, textvariable=self.restaurant_mech_process,
                                       font=("Courier", 9), justify=tk.LEFT, foreground="darkgreen")
        mech_process_label.pack(fill=tk.BOTH, expand=True, anchor=tk.W)

        # ④電気設備フレーム
        elec_frame = ttk.LabelFrame(main_frame, text="電気設備", padding="15")
        elec_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
        row += 1

        self.add_lounge_item(elec_frame, 0, "LED照明", self.restaurant_elec_led_check,
                              self.restaurant_elec_led_count, "個", self.restaurant_elec_led_cost,
                              self.restaurant_elec_led_recommended)
        self.add_lounge_item(elec_frame, 1, "煙感知器", self.restaurant_elec_smoke_check,
                              self.restaurant_elec_smoke_count, "個", self.restaurant_elec_smoke_cost,
                              self.restaurant_elec_smoke_recommended)
        self.add_lounge_item(elec_frame, 2, "誘導灯", self.restaurant_elec_exit_light_check,
                              self.restaurant_elec_exit_light_count, "個", self.restaurant_elec_exit_light_cost,
                              self.restaurant_elec_exit_light_recommended)
        self.add_lounge_item(elec_frame, 3, "スピーカー", self.restaurant_elec_speaker_check,
                              self.restaurant_elec_speaker_count, "個", self.restaurant_elec_speaker_cost,
                              self.restaurant_elec_speaker_recommended)
        self.add_lounge_item(elec_frame, 4, "コンセント", self.restaurant_elec_outlet_check,
                              self.restaurant_elec_outlet_count, "個", self.restaurant_elec_outlet_cost,
                              self.restaurant_elec_outlet_recommended)

        # 電気設備の計算過程
        elec_process_frame = ttk.LabelFrame(main_frame, text="電気設備の計算過程", padding="10")
        elec_process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        elec_process_label = ttk.Label(elec_process_frame, textvariable=self.restaurant_elec_process,
                                       font=("Courier", 9), justify=tk.LEFT, foreground="darkgreen")
        elec_process_label.pack(fill=tk.BOTH, expand=True, anchor=tk.W)

        # ⑤厨房設備フレーム
        kitchen_frame = ttk.LabelFrame(main_frame, text="厨房設備", padding="15")
        kitchen_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
        row += 1

        # 厨房機器設定からデータを動的に生成
        kitchen_row = 0
        for data_vars in self.kitchen_equipment_unit_data:
            no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, expense_var, list_price_var, rate_var, unit_input_var, category_var = data_vars

            # チェックボックス、数量、金額の変数を作成
            check_var = tk.BooleanVar(value=False)
            count_var = tk.StringVar(value="1")
            cost_var = tk.StringVar(value="0")

            self.restaurant_kitchen_equipment_checks.append(check_var)
            self.restaurant_kitchen_equipment_counts.append(count_var)
            self.restaurant_kitchen_equipment_costs.append(cost_var)

            # UIの作成
            self.add_lounge_item(kitchen_frame, kitchen_row, abbrev_var.get(), check_var,
                                 count_var, "台", cost_var)
            kitchen_row += 1
        # 厨房設備の計算過程
        kitchen_process_frame = ttk.LabelFrame(main_frame, text="厨房設備の計算過程", padding="10")
        kitchen_process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        kitchen_process_label = ttk.Label(kitchen_process_frame, textvariable=self.restaurant_kitchen_process,
                                          font=("Courier", 9), justify=tk.LEFT, foreground="purple")
        kitchen_process_label.pack(fill=tk.BOTH, expand=True, anchor=tk.W)


        # 家具フレーム
        furniture_frame = ttk.LabelFrame(main_frame, text="家具", padding="15")
        furniture_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
        row += 1

        self.add_lounge_item(furniture_frame, 0, "バーカウンター", self.restaurant_furniture_bar_counter_check,
                             self.restaurant_furniture_bar_counter_count, "台",
                             self.restaurant_furniture_bar_counter_cost,
                             self.restaurant_furniture_bar_counter_recommended)
        self.add_lounge_item(furniture_frame, 1, "ソフトドリンクカウンター",
                             self.restaurant_furniture_soft_drink_counter_check,
                             self.restaurant_furniture_soft_drink_counter_count, "台",
                             self.restaurant_furniture_soft_drink_counter_cost,
                             self.restaurant_furniture_soft_drink_counter_recommended)
        self.add_lounge_item(furniture_frame, 2, "アルコールカウンター",
                             self.restaurant_furniture_alcohol_counter_check,
                             self.restaurant_furniture_alcohol_counter_count, "台",
                             self.restaurant_furniture_alcohol_counter_cost,
                             self.restaurant_furniture_alcohol_counter_recommended)
        self.add_lounge_item(furniture_frame, 3, "カトラリ―カウンター", self.restaurant_furniture_cutlery_counter_check,
                             self.restaurant_furniture_cutlery_counter_count, "台",
                             self.restaurant_furniture_cutlery_counter_cost,
                             self.restaurant_furniture_cutlery_counter_recommended)
        self.add_lounge_item(furniture_frame, 4, "アイスカウンター", self.restaurant_furniture_ice_counter_check,
                             self.restaurant_furniture_ice_counter_count, "台",
                             self.restaurant_furniture_ice_counter_cost,
                             self.restaurant_furniture_ice_counter_recommended)
        self.add_lounge_item(furniture_frame, 5, "ソフトクリームカウンター",
                             self.restaurant_furniture_soft_cream_counter_check,
                             self.restaurant_furniture_soft_cream_counter_count, "台",
                             self.restaurant_furniture_soft_cream_counter_cost,
                             self.restaurant_furniture_soft_cream_counter_recommended)
        self.add_lounge_item(furniture_frame, 6, "返却台", self.restaurant_furniture_return_counter_check,
                             self.restaurant_furniture_return_counter_count, "台",
                             self.restaurant_furniture_return_counter_cost,
                             self.restaurant_furniture_return_counter_recommended)

        # 家具の計算過程
        furniture_process_frame = ttk.LabelFrame(main_frame, text="家具の計算過程", padding="10")
        furniture_process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        furniture_process_label = ttk.Label(furniture_process_frame, textvariable=self.restaurant_furniture_process,
                                            font=("Courier", 9), justify=tk.LEFT, foreground="purple")
        furniture_process_label.pack(fill=tk.BOTH, expand=True, anchor=tk.W)




        # トレース設定(面積変更時の自動計算)
        self.restaurant_arch_restaurant_area.trace('w', self.update_restaurant_arch_costs)
        self.restaurant_arch_livekitchen_area.trace('w', self.update_restaurant_arch_costs)
        self.restaurant_arch_kitchen_area.trace('w', self.update_restaurant_arch_costs)

        # トレース設定(機械設備・電気設備の個数変更時)
        self.restaurant_mech_ac_count.trace('w', self.update_restaurant_equipment_costs)
        self.restaurant_mech_sprinkler_count.trace('w', self.update_restaurant_equipment_costs)
        self.restaurant_mech_fire_hose_count.trace('w', self.update_restaurant_equipment_costs)
        self.restaurant_mech_hood_count.trace('w', self.update_restaurant_equipment_costs)
        self.restaurant_elec_led_count.trace('w', self.update_restaurant_equipment_costs)
        self.restaurant_elec_smoke_count.trace('w', self.update_restaurant_equipment_costs)
        self.restaurant_elec_exit_light_count.trace('w', self.update_restaurant_equipment_costs)
        self.restaurant_elec_speaker_count.trace('w', self.update_restaurant_equipment_costs)
        self.restaurant_elec_outlet_count.trace('w', self.update_restaurant_equipment_costs)
        # 厨房機器のトレース設定
        for count_var in self.restaurant_kitchen_equipment_counts:
            count_var.trace('w', self.update_restaurant_kitchen_costs)
        # 家具のトレース設定
        self.restaurant_furniture_bar_counter_count.trace('w', self.update_restaurant_furniture_costs)
        self.restaurant_furniture_soft_drink_counter_count.trace('w', self.update_restaurant_furniture_costs)
        self.restaurant_furniture_alcohol_counter_count.trace('w', self.update_restaurant_furniture_costs)
        self.restaurant_furniture_cutlery_counter_count.trace('w', self.update_restaurant_furniture_costs)
        self.restaurant_furniture_ice_counter_count.trace('w', self.update_restaurant_furniture_costs)
        self.restaurant_furniture_soft_cream_counter_count.trace('w', self.update_restaurant_furniture_costs)
        self.restaurant_furniture_return_counter_count.trace('w', self.update_restaurant_furniture_costs)

        # チェックボックス変更時の小計更新
        self.restaurant_arch_restaurant_check.trace('w', self.update_restaurant_subtotal)
        self.restaurant_arch_livekitchen_check.trace('w', self.update_restaurant_subtotal)
        self.restaurant_arch_kitchen_check.trace('w', self.update_restaurant_subtotal)
        self.restaurant_mech_ac_check.trace('w', self.update_restaurant_subtotal)
        self.restaurant_mech_sprinkler_check.trace('w', self.update_restaurant_subtotal)
        self.restaurant_mech_fire_hose_check.trace('w', self.update_restaurant_subtotal)
        self.restaurant_mech_hood_check.trace('w', self.update_restaurant_subtotal)
        self.restaurant_elec_led_check.trace('w', self.update_restaurant_subtotal)
        self.restaurant_elec_smoke_check.trace('w', self.update_restaurant_subtotal)
        self.restaurant_elec_exit_light_check.trace('w', self.update_restaurant_subtotal)
        self.restaurant_elec_speaker_check.trace('w', self.update_restaurant_subtotal)
        self.restaurant_elec_outlet_check.trace('w', self.update_restaurant_subtotal)
        # 厨房機器のチェックボックストレース設定
        for check_var in self.restaurant_kitchen_equipment_checks:
            check_var.trace('w', self.update_restaurant_subtotal)
        self.restaurant_furniture_bar_counter_check.trace('w', self.update_restaurant_subtotal)
        self.restaurant_furniture_soft_drink_counter_check.trace('w', self.update_restaurant_subtotal)
        self.restaurant_furniture_alcohol_counter_check.trace('w', self.update_restaurant_subtotal)
        self.restaurant_furniture_cutlery_counter_check.trace('w', self.update_restaurant_subtotal)
        self.restaurant_furniture_ice_counter_check.trace('w', self.update_restaurant_subtotal)
        self.restaurant_furniture_soft_cream_counter_check.trace('w', self.update_restaurant_furniture_costs)
        self.restaurant_furniture_return_counter_check.trace('w', self.update_restaurant_subtotal)

        # 初期計算
        self.update_restaurant_arch_costs()
        self.update_restaurant_recommended_counts()
        self.update_restaurant_equipment_costs()

    def update_restaurant_arch_costs(self, *args):
        """レストランタブの建築項目の金額を自動計算"""
        # レストラン
        try:
            area = float(self.restaurant_arch_restaurant_area.get())
            if area > 0:
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("レストラン")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    self.restaurant_arch_restaurant_cost.set(f"{int(cost):,}")

                    process = (f"【レストラン - {grade_name}】\n"
                               f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                               f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                               f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                               f"       = {int(cost):,}円")
                    self.restaurant_arch_restaurant_process.set(process)
                else:
                    self.restaurant_arch_restaurant_cost.set("0")
                    self.restaurant_arch_restaurant_process.set("【レストラン】単価情報が取得できません")
            else:
                self.restaurant_arch_restaurant_cost.set("0")
                self.restaurant_arch_restaurant_process.set("")
        except ValueError:
            self.restaurant_arch_restaurant_cost.set("0")
            self.restaurant_arch_restaurant_process.set("")

        # ライブキッチン
        try:
            area = float(self.restaurant_arch_livekitchen_area.get())
            if area > 0:
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("ライブキッチン")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    self.restaurant_arch_livekitchen_cost.set(f"{int(cost):,}")

                    process = (f"【ライブキッチン - {grade_name}】\n"
                               f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                               f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                               f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                               f"       = {int(cost):,}円")
                    self.restaurant_arch_livekitchen_process.set(process)
                else:
                    self.restaurant_arch_livekitchen_cost.set("0")
                    self.restaurant_arch_livekitchen_process.set("【ライブキッチン】単価情報が取得できません")
            else:
                self.restaurant_arch_livekitchen_cost.set("0")
                self.restaurant_arch_livekitchen_process.set("")
        except ValueError:
            self.restaurant_arch_livekitchen_cost.set("0")
            self.restaurant_arch_livekitchen_process.set("")

        # 厨房
        try:
            area = float(self.restaurant_arch_kitchen_area.get())
            if area > 0:
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("厨房")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    self.restaurant_arch_kitchen_cost.set(f"{int(cost):,}")

                    process = (f"【厨房 - {grade_name}】\n"
                               f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                               f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                               f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                               f"       = {int(cost):,}円")
                    self.restaurant_arch_kitchen_process.set(process)
                else:
                    self.restaurant_arch_kitchen_cost.set("0")
                    self.restaurant_arch_kitchen_process.set("【厨房】単価情報が取得できません")
            else:
                self.restaurant_arch_kitchen_cost.set("0")
                self.restaurant_arch_kitchen_process.set("")
        except ValueError:
            self.restaurant_arch_kitchen_cost.set("0")
            self.restaurant_arch_kitchen_process.set("")

        self.update_restaurant_recommended_counts()
        self.update_restaurant_subtotal()

    def update_restaurant_recommended_counts(self, *args):
        """レストランの機械設備・電気設備の推奨個数を計算"""
        try:
            restaurant_area = float(
                self.restaurant_arch_restaurant_area.get()) if self.restaurant_arch_restaurant_area.get() else 0
            livekitchen_area = float(
                self.restaurant_arch_livekitchen_area.get()) if self.restaurant_arch_livekitchen_area.get() else 0
            kitchen_area = float(
                self.restaurant_arch_kitchen_area.get()) if self.restaurant_arch_kitchen_area.get() else 0
            total_area = restaurant_area + livekitchen_area + kitchen_area

            kitchen_area_for_hood = livekitchen_area + kitchen_area#フード個数算定する為厨房面積のみ

            if total_area > 0:
                # 共用部エアコン
                ac_data = self.get_equipment_data_by_abbrev("共用部エアコン")
                if ac_data and ac_data['area_ratio'] > 0:
                    recommended_ac = math.ceil(total_area * ac_data['area_ratio'])
                    self.restaurant_mech_ac_recommended.set(f"推奨: {recommended_ac}台")
                else:
                    self.restaurant_mech_ac_recommended.set("")

                # スプリンクラー
                sprinkler_data = self.get_equipment_data_by_abbrev("スプリンクラー")
                if sprinkler_data and sprinkler_data['area_ratio'] > 0:
                    recommended_sprinkler = math.ceil(total_area * sprinkler_data['area_ratio'])
                    self.restaurant_mech_sprinkler_recommended.set(f"推奨: {recommended_sprinkler}個")
                else:
                    self.restaurant_mech_sprinkler_recommended.set("")

                # 消火栓箱
                fire_hose_data = self.get_equipment_data_by_abbrev("消火栓箱")
                if fire_hose_data and fire_hose_data['area_ratio'] > 0:
                    recommended_fire_hose = math.ceil(total_area * fire_hose_data['area_ratio'])
                    self.restaurant_mech_fire_hose_recommended.set(f"推奨: {recommended_fire_hose}台")
                else:
                    self.restaurant_mech_fire_hose_recommended.set("")

                # フード（レストラン面積を除外してライブキッチン+厨房のみで計算）
                if kitchen_area_for_hood > 0:
                    hood_data = self.get_equipment_data_by_abbrev("フード")
                    if hood_data and hood_data['area_ratio'] > 0:
                        recommended_hood = math.ceil(kitchen_area_for_hood * hood_data['area_ratio'])
                        self.restaurant_mech_hood_recommended.set(f"推奨: {recommended_hood}台")
                    else:
                        self.restaurant_mech_hood_recommended.set("")
                else:
                    self.restaurant_mech_hood_recommended.set("")

                # LED照明(レストランLED)
                led_data = self.get_equipment_data_by_abbrev("レストランLED")
                if led_data and led_data['area_ratio'] > 0:
                    recommended_led = math.ceil(total_area * led_data['area_ratio'])
                    self.restaurant_elec_led_recommended.set(f"推奨: {recommended_led}個")
                else:
                    self.restaurant_elec_led_recommended.set("")

                # 煙感知器
                smoke_data = self.get_equipment_data_by_abbrev("煙感知器")
                if smoke_data and smoke_data['area_ratio'] > 0:
                    recommended_smoke = math.ceil(total_area * smoke_data['area_ratio'])
                    self.restaurant_elec_smoke_recommended.set(f"推奨: {recommended_smoke}個")
                else:
                    self.restaurant_elec_smoke_recommended.set("")

                # 誘導灯
                exit_light_data = self.get_equipment_data_by_abbrev("誘導灯")
                if exit_light_data and exit_light_data['area_ratio'] > 0:
                    recommended_exit_light = math.ceil(total_area * exit_light_data['area_ratio'])
                    self.restaurant_elec_exit_light_recommended.set(f"推奨: {recommended_exit_light}個")
                else:
                    self.restaurant_elec_exit_light_recommended.set("")

                # スピーカー
                speaker_data = self.get_equipment_data_by_abbrev("スピーカー")
                if speaker_data and speaker_data['area_ratio'] > 0:
                    recommended_speaker = math.ceil(total_area * speaker_data['area_ratio'])
                    self.restaurant_elec_speaker_recommended.set(f"推奨: {recommended_speaker}個")
                else:
                    self.restaurant_elec_speaker_recommended.set("")

                # コンセント
                outlet_data = self.get_equipment_data_by_abbrev("コンセント")
                if outlet_data and outlet_data['area_ratio'] > 0:
                    recommended_outlet = math.ceil(total_area * outlet_data['area_ratio'])
                    self.restaurant_elec_outlet_recommended.set(f"推奨: {recommended_outlet}個")
                else:
                    self.restaurant_elec_outlet_recommended.set("")

                # バーカウンター(略称「バーカウンター」で家具タブに登録されていると仮定)
                bar_counter_data = self.get_furniture1_data_by_abbrev("バーカウンター")
                if bar_counter_data and bar_counter_data['area_ratio'] > 0:
                    recommended_bar = math.ceil(total_area * bar_counter_data['area_ratio'])
                    self.restaurant_furniture_bar_counter_recommended.set(f"推奨: {recommended_bar}台")
                else:
                    self.restaurant_furniture_bar_counter_recommended.set("")

                # ソフトドリンクカウンター
                soft_drink_data = self.get_furniture1_data_by_abbrev("ソフトドリンクカウンター")
                if soft_drink_data and soft_drink_data['area_ratio'] > 0:
                    recommended_soft_drink = math.ceil(total_area * soft_drink_data['area_ratio'])
                    self.restaurant_furniture_soft_drink_counter_recommended.set(f"推奨: {recommended_soft_drink}台")
                else:
                    self.restaurant_furniture_soft_drink_counter_recommended.set("")

                # アルコールカウンター
                alcohol_data = self.get_furniture1_data_by_abbrev("アルコールカウンター")
                if alcohol_data and alcohol_data['area_ratio'] > 0:
                    recommended_alcohol = math.ceil(total_area * alcohol_data['area_ratio'])
                    self.restaurant_furniture_alcohol_counter_recommended.set(f"推奨: {recommended_alcohol}台")
                else:
                    self.restaurant_furniture_alcohol_counter_recommended.set("")

                # カトラリ―カウンター
                cutlery_data = self.get_furniture1_data_by_abbrev("カトラリ―カウンター")
                if cutlery_data and cutlery_data['area_ratio'] > 0:
                    recommended_cutlery = math.ceil(total_area * cutlery_data['area_ratio'])
                    self.restaurant_furniture_cutlery_counter_recommended.set(f"推奨: {recommended_cutlery}台")
                else:
                    self.restaurant_furniture_cutlery_counter_recommended.set("")

                # アイスカウンター
                ice_data = self.get_furniture1_data_by_abbrev("アイスカウンター")
                if ice_data and ice_data['area_ratio'] > 0:
                    recommended_ice = math.ceil(total_area * ice_data['area_ratio'])
                    self.restaurant_furniture_ice_counter_recommended.set(f"推奨: {recommended_ice}台")
                else:
                    self.restaurant_furniture_ice_counter_recommended.set("")

                # ソフトクリームカウンター
                soft_cream_data = self.get_furniture1_data_by_abbrev("ソフトクリームカウンター")
                if soft_cream_data and soft_cream_data['area_ratio'] > 0:
                    recommended_soft_cream = math.ceil(total_area * soft_cream_data['area_ratio'])
                    self.restaurant_furniture_soft_cream_counter_recommended.set(f"推奨: {recommended_soft_cream}台")
                else:
                    self.restaurant_furniture_soft_cream_counter_recommended.set("")

                # 返却台
                return_data = self.get_furniture1_data_by_abbrev("返却台")
                if return_data and return_data['area_ratio'] > 0:
                    recommended_return = math.ceil(total_area * return_data['area_ratio'])
                    self.restaurant_furniture_return_counter_recommended.set(f"推奨: {recommended_return}台")
                else:
                    self.restaurant_furniture_return_counter_recommended.set("")

            else:
                self.restaurant_mech_ac_recommended.set("")
                self.restaurant_mech_sprinkler_recommended.set("")
                self.restaurant_mech_fire_hose_recommended.set("")
                self.restaurant_mech_hood_recommended.set("")
                self.restaurant_elec_led_recommended.set("")
                self.restaurant_elec_smoke_recommended.set("")
                self.restaurant_elec_exit_light_recommended.set("")
                self.restaurant_elec_speaker_recommended.set("")
                self.restaurant_elec_outlet_recommended.set("")

        except (ValueError, AttributeError):
            pass

    def update_restaurant_equipment_costs(self, *args):
        """レストランの機械設備・電気設備の金額を計算"""
        mech_process_text = "【機械設備の計算】\n"
        elec_process_text = "【電気設備の計算】\n"

        # 共用部エアコン
        try:
            count = float(self.restaurant_mech_ac_count.get()) if self.restaurant_mech_ac_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("共用部エアコン")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.restaurant_mech_ac_cost.set(f"{int(cost):,}")
                    mech_process_text += f"\n[共用部エアコン]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.restaurant_mech_ac_cost.set("0")
        except (ValueError, TypeError):
            self.restaurant_mech_ac_cost.set("0")

        # スプリンクラー
        try:
            count = float(
                self.restaurant_mech_sprinkler_count.get()) if self.restaurant_mech_sprinkler_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("スプリンクラー")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.restaurant_mech_sprinkler_cost.set(f"{int(cost):,}")
                    mech_process_text += f"\n[スプリンクラー]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.restaurant_mech_sprinkler_cost.set("0")
        except (ValueError, TypeError):
            self.restaurant_mech_sprinkler_cost.set("0")

        # 消火栓箱
        try:
            count = float(
                self.restaurant_mech_fire_hose_count.get()) if self.restaurant_mech_fire_hose_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("消火栓箱")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.restaurant_mech_fire_hose_cost.set(f"{int(cost):,}")
                    mech_process_text += f"\n[消火栓箱]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.restaurant_mech_fire_hose_cost.set("0")
        except (ValueError, TypeError):
            self.restaurant_mech_fire_hose_cost.set("0")

        # フード
        try:
            count = float(
                self.restaurant_mech_hood_count.get()) if self.restaurant_mech_hood_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("フード")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.restaurant_mech_hood_cost.set(f"{int(cost):,}")
                    mech_process_text += f"\n[フード]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
                else:
                    self.restaurant_mech_hood_cost.set("0")
                    mech_process_text += f"\n[フード]\n  ※設備機器着脱単価設定に「フード」を追加してください\n"
            else:
                self.restaurant_mech_hood_cost.set("0")
        except (ValueError, TypeError):
            self.restaurant_mech_hood_cost.set("0")

        # LED照明
        try:
            count = float(self.restaurant_elec_led_count.get()) if self.restaurant_elec_led_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("レストランLED")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.restaurant_elec_led_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[LED照明]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.restaurant_elec_led_cost.set("0")
        except (ValueError, TypeError):
            self.restaurant_elec_led_cost.set("0")

        # 煙感知器
        try:
            count = float(self.restaurant_elec_smoke_count.get()) if self.restaurant_elec_smoke_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("煙感知器")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.restaurant_elec_smoke_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[煙感知器]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.restaurant_elec_smoke_cost.set("0")
        except (ValueError, TypeError):
            self.restaurant_elec_smoke_cost.set("0")

        # 誘導灯
        try:
            count = float(
                self.restaurant_elec_exit_light_count.get()) if self.restaurant_elec_exit_light_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("誘導灯")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.restaurant_elec_exit_light_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[誘導灯]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.restaurant_elec_exit_light_cost.set("0")
        except (ValueError, TypeError):
            self.restaurant_elec_exit_light_cost.set("0")

        # スピーカー
        try:
            count = float(self.restaurant_elec_speaker_count.get()) if self.restaurant_elec_speaker_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("スピーカー")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.restaurant_elec_speaker_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[スピーカー]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.restaurant_elec_speaker_cost.set("0")
        except (ValueError, TypeError):
            self.restaurant_elec_speaker_cost.set("0")

        # コンセント
        try:
            count = float(self.restaurant_elec_outlet_count.get()) if self.restaurant_elec_outlet_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("コンセント")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.restaurant_elec_outlet_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[コンセント]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.restaurant_elec_outlet_cost.set("0")
        except (ValueError, TypeError):
            self.restaurant_elec_outlet_cost.set("0")

        self.restaurant_mech_process.set(mech_process_text)
        self.restaurant_elec_process.set(elec_process_text)

        # 厨房設備の計算
        self.update_restaurant_kitchen_costs()
        # 小計を更新
        self.update_restaurant_subtotal()

    def update_restaurant_kitchen_costs(self, *args):
        """レストランの厨房設備の金額を計算"""
        kitchen_process_text = "【厨房設備の計算】\n"

        for idx, data_vars in enumerate(self.kitchen_equipment_unit_data):
            no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, expense_var, list_price_var, rate_var, unit_input_var, category_var = data_vars

            try:
                count = float(self.restaurant_kitchen_equipment_counts[idx].get()) if \
                self.restaurant_kitchen_equipment_counts[idx].get() else 0
                if count > 0:
                    # 厨房機器の単価計算
                    ih = float(install_hours_var.get()) if install_hours_var.get() else 0
                    lc = float(labor_cost_var.get()) if labor_cost_var.get() else 0
                    exp = float(expense_var.get()) if expense_var.get() else 0
                    lp = float(list_price_var.get()) if list_price_var.get() else 0
                    r = float(rate_var.get()) if rate_var.get() else 0

                    # 機器単価 = (取付工数 × 労務単価) + 経費 + (定価 × 掛け率)
                    unit_cost = (ih * lc) + exp + (lp * r)
                    total_cost = unit_cost * count

                    self.restaurant_kitchen_equipment_costs[idx].set(f"{int(total_cost):,}")

                    kitchen_process_text += (f"\n[{abbrev_var.get()}]\n"
                                             f"  単価 = ({ih}×{lc:,}) + {exp:,} + ({lp:,}×{r})\n"
                                             f"       = {int(unit_cost):,}円\n"
                                             f"  金額 = {int(unit_cost):,} × {count} = {int(total_cost):,}円\n")
                else:
                    self.restaurant_kitchen_equipment_costs[idx].set("0")
            except (ValueError, TypeError, IndexError):
                if idx < len(self.restaurant_kitchen_equipment_costs):
                    self.restaurant_kitchen_equipment_costs[idx].set("0")

        self.restaurant_kitchen_process.set(kitchen_process_text)
        self.update_restaurant_subtotal()

    def update_restaurant_subtotal(self, *args):
        """レストラン小計を計算してリアルタイム更新"""
        total = 0

        # 建築項目
        if self.restaurant_arch_restaurant_check.get():
            try:
                cost = float(self.restaurant_arch_restaurant_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.restaurant_arch_livekitchen_check.get():
            try:
                cost = float(self.restaurant_arch_livekitchen_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.restaurant_arch_kitchen_check.get():
            try:
                cost = float(self.restaurant_arch_kitchen_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # 機械設備項目
        if self.restaurant_mech_ac_check.get():
            try:
                cost = float(self.restaurant_mech_ac_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.restaurant_mech_sprinkler_check.get():
            try:
                cost = float(self.restaurant_mech_sprinkler_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.restaurant_mech_fire_hose_check.get():
            try:
                cost = float(self.restaurant_mech_fire_hose_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.restaurant_mech_hood_check.get():
            try:
                cost = float(self.restaurant_mech_hood_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # 電気設備項目
        if self.restaurant_elec_led_check.get():
            try:
                cost = float(self.restaurant_elec_led_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.restaurant_elec_smoke_check.get():
            try:
                cost = float(self.restaurant_elec_smoke_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.restaurant_elec_exit_light_check.get():
            try:
                cost = float(self.restaurant_elec_exit_light_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.restaurant_elec_speaker_check.get():
            try:
                cost = float(self.restaurant_elec_speaker_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.restaurant_elec_outlet_check.get():
            try:
                cost = float(self.restaurant_elec_outlet_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass
        # 厨房設備項目
        for idx, check_var in enumerate(self.restaurant_kitchen_equipment_checks):
            if check_var.get():
                try:
                    cost = float(self.restaurant_kitchen_equipment_costs[idx].get().replace(',', ''))
                    total += cost
                except (ValueError, AttributeError, IndexError):
                    pass
        # 家具項目
        if self.restaurant_furniture_bar_counter_check.get():
            try:
                cost = float(self.restaurant_furniture_bar_counter_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.restaurant_furniture_soft_drink_counter_check.get():
            try:
                cost = float(self.restaurant_furniture_soft_drink_counter_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.restaurant_furniture_alcohol_counter_check.get():
            try:
                cost = float(self.restaurant_furniture_alcohol_counter_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.restaurant_furniture_cutlery_counter_check.get():
            try:
                cost = float(self.restaurant_furniture_cutlery_counter_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.restaurant_furniture_ice_counter_check.get():
            try:
                cost = float(self.restaurant_furniture_ice_counter_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.restaurant_furniture_soft_cream_counter_check.get():
            try:
                cost = float(self.restaurant_furniture_soft_cream_counter_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.restaurant_furniture_return_counter_check.get():
            try:
                cost = float(self.restaurant_furniture_return_counter_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # 小計を更新
        self.restaurant_subtotal.set(f" {int(total):,}")
    # レストランのタブに関する項目終了

    # ラウンジのタブに関する項目開始
    def add_lounge_item(self, frame, row, label_text, check_var, input_var, unit_text, cost_var, recommended_var=None):
        """ラウンジタブの項目を作成するヘルパー関数"""
        ttk.Checkbutton(frame, text=label_text, variable=check_var).grid(row=row, column=0, sticky=tk.W, padx=(0, 10),
                                                                         pady=3)
        ttk.Entry(frame, textvariable=input_var, width=10, justify=tk.RIGHT).grid(row=row, column=1, sticky=tk.W,
                                                                                  padx=(5, 5), pady=3)
        ttk.Label(frame, text=unit_text).grid(row=row, column=2, sticky=tk.W, padx=(5, 10), pady=3)

        if recommended_var is not None:
            ttk.Label(frame, textvariable=recommended_var, font=("Arial", 9, "bold"), foreground="blue").grid(row=row,
                                                                                                              column=3,
                                                                                                              sticky=tk.W,
                                                                                                              padx=(5,
                                                                                                                    10),
                                                                                                              pady=3)

        ttk.Entry(frame, textvariable=cost_var, width=12, state='readonly', justify=tk.RIGHT).grid(row=row, column=4,
                                                                                                   sticky=tk.W,
                                                                                                   padx=(5, 5), pady=3)
        ttk.Label(frame, text="円").grid(row=row, column=5, sticky=tk.W, pady=3)

    def create_lounge_tab(self):
        """ラウンジタブの作成"""
        canvas = tk.Canvas(self.lounge_frame)
        scrollbar = ttk.Scrollbar(self.lounge_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(main_frame, text="ラウンジ詳細設定", font=title_font).grid(row=0, column=0, columnspan=3,
                                                                                        pady=(0, 20), sticky=tk.W)
        ttk.Label(main_frame, text="面積を入力すると金額が表示されます。計上金額に含める場合は、項目に☑を入れてください。",
                  font=("Arial", 8)).grid(row=1, column=0, columnspan=3, sticky=tk.E, pady=(20, 0))
        row = 2

        # ラウンジ小計
        subtotal_frame = ttk.Frame(main_frame)
        subtotal_frame.grid(row=row, column=0, sticky=tk.W, pady=8)
        ttk.Label(subtotal_frame, text="小計:", font=item_font).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Entry(subtotal_frame, textvariable=self.lounge_subtotal, width=10, state='readonly',
                  font=item_font, justify='right').pack(side=tk.LEFT)
        ttk.Label(subtotal_frame, text=" 円", font=item_font).pack(side=tk.LEFT, padx=(5, 0))
        row += 1

        arch_frame = ttk.LabelFrame(main_frame, text="建築", padding="15")
        arch_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1

        self.add_lounge_item(arch_frame, 0, "フロント", self.lounge_arch_front_check, self.lounge_arch_front_area, "㎡",
                              self.lounge_arch_front_cost)
        self.add_lounge_item(arch_frame, 1, "ロビー", self.lounge_arch_lobby_check, self.lounge_arch_lobby_area, "㎡",
                              self.lounge_arch_lobby_cost)
        self.add_lounge_item(arch_frame, 2, "ファサード", self.lounge_arch_facade_check, self.lounge_arch_facade_area,
                              "㎡", self.lounge_arch_facade_cost)
        self.add_lounge_item(arch_frame, 3, "売店", self.lounge_arch_shop_check, self.lounge_arch_shop_area,
                              "㎡", self.lounge_arch_shop_cost)

        process_frame = ttk.LabelFrame(main_frame, text="建築項目の計算過程", padding="10")
        process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        process_text_frame = ttk.Frame(process_frame)
        process_text_frame.pack(fill=tk.BOTH, expand=True)

        front_label = ttk.Label(process_text_frame, textvariable=self.lounge_arch_front_process, font=("Courier", 9),
                                justify=tk.LEFT, foreground="blue")
        front_label.pack(anchor=tk.W, pady=3)

        lobby_label = ttk.Label(process_text_frame, textvariable=self.lounge_arch_lobby_process, font=("Courier", 9),
                                justify=tk.LEFT, foreground="blue")
        lobby_label.pack(anchor=tk.W, pady=3)

        facade_label = ttk.Label(process_text_frame, textvariable=self.lounge_arch_facade_process, font=("Courier", 9),
                                 justify=tk.LEFT, foreground="blue")
        facade_label.pack(anchor=tk.W, pady=3)

        shop_label = ttk.Label(process_text_frame, textvariable=self.lounge_arch_shop_process, font=("Courier", 9),
                               justify=tk.LEFT, foreground="blue")
        shop_label.pack(anchor=tk.W, pady=3)

        mech_frame = ttk.LabelFrame(main_frame, text="機械設備", padding="15")
        mech_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
        row += 1

        self.add_lounge_item(mech_frame, 0, "共用部エアコン", self.lounge_mech_ac_check, self.lounge_mech_ac_count,
                              "台", self.lounge_mech_ac_cost, self.lounge_mech_ac_recommended)
        self.add_lounge_item(mech_frame, 1, "スプリンクラー", self.lounge_mech_sprinkler_check,
                              self.lounge_mech_sprinkler_count, "個", self.lounge_mech_sprinkler_cost,
                              self.lounge_mech_sprinkler_recommended)
        self.add_lounge_item(mech_frame, 2, "消火栓箱", self.lounge_mech_fire_hose_check,
                              self.lounge_mech_fire_hose_count, "台", self.lounge_mech_fire_hose_cost,
                              self.lounge_mech_fire_hose_recommended)

        mech_process_frame = ttk.LabelFrame(main_frame, text="機械設備の計算過程", padding="10")
        mech_process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        mech_process_label = ttk.Label(mech_process_frame, textvariable=self.lounge_mech_process, font=("Courier", 9),
                                       justify=tk.LEFT, foreground="darkgreen")
        mech_process_label.pack(fill=tk.BOTH, expand=True, anchor=tk.W)

        elec_frame = ttk.LabelFrame(main_frame, text="電気設備", padding="15")
        elec_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
        row += 1

        self.add_lounge_item(elec_frame, 0, "LED照明", self.lounge_elec_led_check, self.lounge_elec_led_count, "個",
                              self.lounge_elec_led_cost, self.lounge_elec_led_recommended)
        self.add_lounge_item(elec_frame, 1, "煙感知器", self.lounge_elec_smoke_check, self.lounge_elec_smoke_count,
                              "個", self.lounge_elec_smoke_cost, self.lounge_elec_smoke_recommended)
        self.add_lounge_item(elec_frame, 2, "誘導灯", self.lounge_elec_exit_light_check,
                              self.lounge_elec_exit_light_count, "個", self.lounge_elec_exit_light_cost,
                              self.lounge_elec_exit_light_recommended)
        self.add_lounge_item(elec_frame, 3, "スピーカー", self.lounge_elec_speaker_check,
                              self.lounge_elec_speaker_count, "個", self.lounge_elec_speaker_cost,
                              self.lounge_elec_speaker_recommended)
        self.add_lounge_item(elec_frame, 4, "コンセント", self.lounge_elec_outlet_check, self.lounge_elec_outlet_count,
                              "個", self.lounge_elec_outlet_cost, self.lounge_elec_outlet_recommended)

        elec_process_frame = ttk.LabelFrame(main_frame, text="電気設備の計算過程", padding="10")
        elec_process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        elec_process_label = ttk.Label(elec_process_frame, textvariable=self.lounge_elec_process, font=("Courier", 9),
                                       justify=tk.LEFT, foreground="darkgreen")
        elec_process_label.pack(fill=tk.BOTH, expand=True, anchor=tk.W)

    def update_lounge_equipment_costs(self, *args):
        """ラウンジの機械設備・電気設備の金額を計算"""
        mech_process_text = "【機械設備の計算】\n"
        elec_process_text = "【電気設備の計算】\n"

        try:
            count = float(self.lounge_mech_ac_count.get()) if self.lounge_mech_ac_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("共用部エアコン")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.lounge_mech_ac_cost.set(f"{int(cost):,}")
                    mech_process_text += f"\n[共用部エアコン]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.lounge_mech_ac_cost.set("0")
        except (ValueError, TypeError):
            self.lounge_mech_ac_cost.set("0")

        try:
            count = float(self.lounge_mech_sprinkler_count.get()) if self.lounge_mech_sprinkler_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("スプリンクラー")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.lounge_mech_sprinkler_cost.set(f"{int(cost):,}")
                    mech_process_text += f"\n[スプリンクラー]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.lounge_mech_sprinkler_cost.set("0")
        except (ValueError, TypeError):
            self.lounge_mech_sprinkler_cost.set("0")

        try:
            count = float(self.lounge_mech_fire_hose_count.get()) if self.lounge_mech_fire_hose_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("消火栓箱")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.lounge_mech_fire_hose_cost.set(f"{int(cost):,}")
                    mech_process_text += f"\n[消火栓箱]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.lounge_mech_fire_hose_cost.set("0")
        except (ValueError, TypeError):
            self.lounge_mech_fire_hose_cost.set("0")

        try:
            count = float(self.lounge_elec_led_count.get()) if self.lounge_elec_led_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("エントランスLED")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.lounge_elec_led_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[LED照明]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.lounge_elec_led_cost.set("0")
        except (ValueError, TypeError):
            self.lounge_elec_led_cost.set("0")

        try:
            count = float(self.lounge_elec_smoke_count.get()) if self.lounge_elec_smoke_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("煙感知器")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.lounge_elec_smoke_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[煙感知器]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.lounge_elec_smoke_cost.set("0")
        except (ValueError, TypeError):
            self.lounge_elec_smoke_cost.set("0")

        try:
            count = float(self.lounge_elec_exit_light_count.get()) if self.lounge_elec_exit_light_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("誘導灯")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.lounge_elec_exit_light_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[誘導灯]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.lounge_elec_exit_light_cost.set("0")
        except (ValueError, TypeError):
            self.lounge_elec_exit_light_cost.set("0")

        try:
            count = float(self.lounge_elec_speaker_count.get()) if self.lounge_elec_speaker_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("スピーカー")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.lounge_elec_speaker_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[スピーカー]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.lounge_elec_speaker_cost.set("0")
        except (ValueError, TypeError):
            self.lounge_elec_speaker_cost.set("0")

        try:
            count = float(self.lounge_elec_outlet_count.get()) if self.lounge_elec_outlet_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("コンセント")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.lounge_elec_outlet_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[コンセント]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.lounge_elec_outlet_cost.set("0")
        except (ValueError, TypeError):
            self.lounge_elec_outlet_cost.set("0")

        self.lounge_mech_process.set(mech_process_text)
        self.lounge_elec_process.set(elec_process_text)

        # 小計を更新
        self.update_lounge_subtotal()

    def get_lounge_unit_costs(self, room_name):
        """建築初期設定タブから特定の室用途の内装単価と天井高を取得"""
        grade_index = self.get_selected_spec_grade_index()

        for name, height_var, cost_vars in self.arch_unit_costs:
            if name == room_name:
                try:
                    height = float(height_var.get())
                    floor_cost = float(cost_vars[grade_index].get())
                    ceiling_cost = float(cost_vars[5 + grade_index].get())
                    wall_cost = float(cost_vars[10 + grade_index].get())

                    grade_names = ["Gensen", "Premium 1", "Premium 2", "TAOYA 1", "TAOYA 2"]
                    selected_grade = grade_names[grade_index]

                    return height, floor_cost, ceiling_cost, wall_cost, selected_grade
                except (ValueError, IndexError):
                    return None, None, None, None, None
        return None, None, None, None, None

    def update_lounge_recommended_counts(self, *args):
        """ラウンジの機械設備・電気設備の推奨個数を計算"""
        try:
            front_area = float(self.lounge_arch_front_area.get()) if self.lounge_arch_front_area.get() else 0
            lobby_area = float(self.lounge_arch_lobby_area.get()) if self.lounge_arch_lobby_area.get() else 0
            facade_area = float(self.lounge_arch_facade_area.get()) if self.lounge_arch_facade_area.get() else 0
            shop_area = float(self.lounge_arch_shop_area.get()) if self.lounge_arch_shop_area.get() else 0
            total_area = front_area + lobby_area + facade_area + shop_area

            if total_area > 0:
                ac_data = self.get_equipment_data_by_abbrev("共用部エアコン")
                if ac_data and ac_data['area_ratio'] > 0:
                    recommended_ac = math.ceil(total_area * ac_data['area_ratio'])
                    self.lounge_mech_ac_recommended.set(f"推奨: {recommended_ac}台")
                else:
                    self.lounge_mech_ac_recommended.set("")

                sprinkler_data = self.get_equipment_data_by_abbrev("スプリンクラー")
                if sprinkler_data and sprinkler_data['area_ratio'] > 0:
                    recommended_sprinkler = math.ceil(total_area * sprinkler_data['area_ratio'])
                    self.lounge_mech_sprinkler_recommended.set(f"推奨: {recommended_sprinkler}個")
                else:
                    self.lounge_mech_sprinkler_recommended.set("")

                fire_hose_data = self.get_equipment_data_by_abbrev("消火栓箱")
                if fire_hose_data and fire_hose_data['area_ratio'] > 0:
                    recommended_fire_hose = math.ceil(total_area * fire_hose_data['area_ratio'])
                    self.lounge_mech_fire_hose_recommended.set(f"推奨: {recommended_fire_hose}台")
                else:
                    self.lounge_mech_fire_hose_recommended.set("")

                led_found = False
                for abbrev in ["エントランスLED", "通路LED", "レストランLED"]:
                    led_data = self.get_equipment_data_by_abbrev(abbrev)
                    if led_data and led_data['area_ratio'] > 0:
                        recommended_led = math.ceil(total_area * led_data['area_ratio'])
                        self.lounge_elec_led_recommended.set(f"推奨: {recommended_led}個")
                        led_found = True
                        break
                if not led_found:
                    self.lounge_elec_led_recommended.set("")

                smoke_data = self.get_equipment_data_by_abbsmoke_data = self.get_equipment_data_by_abbrev("煙感知器")
                if smoke_data and smoke_data['area_ratio'] > 0:
                    recommended_smoke = math.ceil(total_area * smoke_data['area_ratio'])
                    self.lounge_elec_smoke_recommended.set(f"推奨: {recommended_smoke}個")
                else:
                    self.lounge_elec_smoke_recommended.set("")

                exit_light_data = self.get_equipment_data_by_abbrev("誘導灯")
                if exit_light_data and exit_light_data['area_ratio'] > 0:
                    recommended_exit_light = math.ceil(total_area * exit_light_data['area_ratio'])
                    self.lounge_elec_exit_light_recommended.set(f"推奨: {recommended_exit_light}個")
                else:
                    self.lounge_elec_exit_light_recommended.set("")

                speaker_data = self.get_equipment_data_by_abbrev("スピーカー")
                if speaker_data and speaker_data['area_ratio'] > 0:
                    recommended_speaker = math.ceil(total_area * speaker_data['area_ratio'])
                    self.lounge_elec_speaker_recommended.set(f"推奨: {recommended_speaker}個")
                else:
                    self.lounge_elec_speaker_recommended.set("")

                outlet_data = self.get_equipment_data_by_abbrev("コンセント")
                if outlet_data and outlet_data['area_ratio'] > 0:
                    recommended_outlet = math.ceil(total_area * outlet_data['area_ratio'])
                    self.lounge_elec_outlet_recommended.set(f"推奨: {recommended_outlet}個")
                else:
                    self.lounge_elec_outlet_recommended.set("")
            else:
                self.lounge_mech_ac_recommended.set("")
                self.lounge_mech_sprinkler_recommended.set("")
                self.lounge_mech_fire_hose_recommended.set("")
                self.lounge_elec_led_recommended.set("")
                self.lounge_elec_smoke_recommended.set("")
                self.lounge_elec_exit_light_recommended.set("")
                self.lounge_elec_speaker_recommended.set("")
                self.lounge_elec_outlet_recommended.set("")

        except (ValueError, AttributeError):
            pass

    def update_lounge_arch_costs(self, *args):
        """ラウンジタブの建築項目の金額を自動計算"""
        try:
            area = float(self.lounge_arch_front_area.get())
            if area > 0:
                # 初期設定から室用途を照合する検索キーワード
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("フロント")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    self.lounge_arch_front_cost.set(f"{int(cost):,}")

                    process = (f"【フロント - {grade_name}】\n"
                               f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                               f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                               f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                               f"       = {int(cost):,}円")
                    self.lounge_arch_front_process.set(process)
                else:
                    self.lounge_arch_front_cost.set("0")
                    self.lounge_arch_front_process.set("【フロント】単価情報が取得できません")
            else:
                self.lounge_arch_front_cost.set("0")
                self.lounge_arch_front_process.set("")
        except ValueError:
            self.lounge_arch_front_cost.set("0")
            self.lounge_arch_front_process.set("")

        try:
            area = float(self.lounge_arch_lobby_area.get())
            if area > 0:
                # 初期設定から室用途を照合する検索キーワード
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("ロビー")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    self.lounge_arch_lobby_cost.set(f"{int(cost):,}")

                    process = (f"【ロビー - {grade_name}】\n"
                               f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                               f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                               f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                               f"       = {int(cost):,}円")
                    self.lounge_arch_lobby_process.set(process)
                else:
                    self.lounge_arch_lobby_cost.set("0")
                    self.lounge_arch_lobby_process.set("【ロビー】単価情報が取得できません")
            else:
                self.lounge_arch_lobby_cost.set("0")
                self.lounge_arch_lobby_process.set("")
        except ValueError:
            self.lounge_arch_lobby_cost.set("0")
            self.lounge_arch_lobby_process.set("")

        try:
            area = float(self.lounge_arch_facade_area.get())
            if area > 0:
                # 初期設定から室用途を照合する検索キーワード
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("ファサード")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    self.lounge_arch_facade_cost.set(f"{int(cost):,}")

                    process = (f"【ファサード - {grade_name}】\n"
                               f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                               f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                               f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                               f"       = {int(cost):,}円")
                    self.lounge_arch_facade_process.set(process)
                else:
                    self.lounge_arch_facade_cost.set("0")
                    self.lounge_arch_facade_process.set("【ファサード】単価情報が取得できません")
            else:
                self.lounge_arch_facade_cost.set("0")
                self.lounge_arch_facade_process.set("")
        except ValueError:
            self.lounge_arch_facade_cost.set("0")
            self.lounge_arch_facade_process.set("")

        try:
            area = float(self.lounge_arch_shop_area.get())
            if area > 0:
                # 建築初期設定から室用途を照合する検索キーワード
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("売店")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    self.lounge_arch_shop_cost.set(f"{int(cost):,}")

                    process = (f"【売店 - {grade_name}】\n"
                               f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                               f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                               f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                               f"       = {int(cost):,}円")
                    self.lounge_arch_shop_process.set(process)
                else:
                    self.lounge_arch_shop_cost.set("0")
                    self.lounge_arch_shop_process.set("【売店】単価情報が取得できません")
            else:
                self.lounge_arch_shop_cost.set("0")
                self.lounge_arch_shop_process.set("")
        except ValueError:
            self.lounge_arch_shop_cost.set("0")
            self.lounge_arch_shop_process.set("")

        self.update_lounge_recommended_counts()
        self.update_lounge_subtotal()

    def update_lounge_subtotal(self, *args):
        """ラウンジ小計を計算してリアルタイム更新"""
        total = 0

        # 建築項目
        if self.lounge_arch_front_check.get():
            try:
                cost = float(self.lounge_arch_front_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.lounge_arch_lobby_check.get():
            try:
                cost = float(self.lounge_arch_lobby_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.lounge_arch_facade_check.get():
            try:
                cost = float(self.lounge_arch_facade_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.lounge_arch_shop_check.get():
            try:
                cost = float(self.lounge_arch_shop_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # 機械設備項目
        if self.lounge_mech_ac_check.get():
            try:
                cost = float(self.lounge_mech_ac_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.lounge_mech_sprinkler_check.get():
            try:
                cost = float(self.lounge_mech_sprinkler_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.lounge_mech_fire_hose_check.get():
            try:
                cost = float(self.lounge_mech_fire_hose_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # 電気設備項目
        if self.lounge_elec_led_check.get():
            try:
                cost = float(self.lounge_elec_led_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.lounge_elec_smoke_check.get():
            try:
                cost = float(self.lounge_elec_smoke_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.lounge_elec_exit_light_check.get():
            try:
                cost = float(self.lounge_elec_exit_light_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.lounge_elec_speaker_check.get():
            try:
                cost = float(self.lounge_elec_speaker_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.lounge_elec_outlet_check.get():
            try:
                cost = float(self.lounge_elec_outlet_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # 小計を更新
        self.lounge_subtotal.set(f" {int(total):,}")
    # ラウンジのタブに関する項目終了

    # 通路のタブに関する項目開始
    def create_hallway_tab(self):
        """通路タブの作成"""
        canvas = tk.Canvas(self.hallway_frame)
        scrollbar = ttk.Scrollbar(self.hallway_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(main_frame, text="通路詳細設定", font=title_font).grid(row=0, column=0, columnspan=3,
                                                                                    pady=(0, 20), sticky=tk.W)
        ttk.Label(main_frame, text="面積を入力すると金額が表示されます。計上金額に含める場合は、項目に☑を入れてください。",
                  font=("Arial", 8)).grid(row=1, column=0, columnspan=3, sticky=tk.E, pady=(20, 0))
        row = 2

        # 通路小計
        subtotal_frame = ttk.Frame(main_frame)
        subtotal_frame.grid(row=row, column=0, sticky=tk.W, pady=8)
        ttk.Label(subtotal_frame, text="小計:", font=item_font).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Entry(subtotal_frame, textvariable=self.hallway_subtotal, width=10, state='readonly',
                  font=item_font, justify='right').pack(side=tk.LEFT)
        ttk.Label(subtotal_frame, text=" 円", font=item_font).pack(side=tk.LEFT, padx=(5, 0))
        row += 1

        # 建築フレーム（動的に生成される各階通路）
        self.hallway_arch_frame = ttk.LabelFrame(main_frame, text="建築", padding="15")
        self.hallway_arch_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1

        # 建築項目の計算過程フレーム
        self.hallway_arch_process_frame = ttk.LabelFrame(main_frame, text="建築項目の計算過程", padding="10")
        self.hallway_arch_process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        # 機械設備
        mech_frame = ttk.LabelFrame(main_frame, text="機械設備", padding="15")
        mech_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
        row += 1

        self.add_lounge_item(mech_frame, 0, "共用部エアコン", self.hallway_mech_ac_check, self.hallway_mech_ac_count,
                              "台", self.hallway_mech_ac_cost, self.hallway_mech_ac_recommended)
        self.add_lounge_item(mech_frame, 1, "スプリンクラー", self.hallway_mech_sprinkler_check,
                              self.hallway_mech_sprinkler_count, "個", self.hallway_mech_sprinkler_cost,
                              self.hallway_mech_sprinkler_recommended)
        self.add_lounge_item(mech_frame, 2, "消火栓箱", self.hallway_mech_fire_hose_check,
                              self.hallway_mech_fire_hose_count, "台", self.hallway_mech_fire_hose_cost,
                              self.hallway_mech_fire_hose_recommended)

        # 機械設備の計算過程
        mech_process_frame = ttk.LabelFrame(main_frame, text="機械設備の計算過程", padding="10")
        mech_process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        mech_process_label = ttk.Label(mech_process_frame, textvariable=self.hallway_mech_process, font=("Courier", 9),
                                       justify=tk.LEFT, foreground="darkgreen")
        mech_process_label.pack(fill=tk.BOTH, expand=True, anchor=tk.W)

        # 電気設備
        elec_frame = ttk.LabelFrame(main_frame, text="電気設備", padding="15")
        elec_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
        row += 1

        self.add_lounge_item(elec_frame, 0, "LED照明", self.hallway_elec_led_check, self.hallway_elec_led_count, "個",
                              self.hallway_elec_led_cost, self.hallway_elec_led_recommended)
        self.add_lounge_item(elec_frame, 1, "煙感知器", self.hallway_elec_smoke_check, self.hallway_elec_smoke_count,
                              "個", self.hallway_elec_smoke_cost, self.hallway_elec_smoke_recommended)
        self.add_lounge_item(elec_frame, 2, "誘導灯", self.hallway_elec_exit_light_check,
                              self.hallway_elec_exit_light_count, "個", self.hallway_elec_exit_light_cost,
                              self.hallway_elec_exit_light_recommended)
        self.add_lounge_item(elec_frame, 3, "スピーカー", self.hallway_elec_speaker_check,
                              self.hallway_elec_speaker_count, "個", self.hallway_elec_speaker_cost,
                              self.hallway_elec_speaker_recommended)
        self.add_lounge_item(elec_frame, 4, "コンセント", self.hallway_elec_outlet_check,
                              self.hallway_elec_outlet_count,
                              "個", self.hallway_elec_outlet_cost, self.hallway_elec_outlet_recommended)

        # 電気設備の計算過程
        elec_process_frame = ttk.LabelFrame(main_frame, text="電気設備の計算過程", padding="10")
        elec_process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        elec_process_label = ttk.Label(elec_process_frame, textvariable=self.hallway_elec_process, font=("Courier", 9),
                                       justify=tk.LEFT, foreground="darkgreen")
        elec_process_label.pack(fill=tk.BOTH, expand=True, anchor=tk.W)

    def update_hallway_floors(self, *args):
        """階数に応じて通路タブの建築項目を動的に生成"""
        try:
            above_floors = int(self.floors.get()) if self.floors.get() else 0
        except ValueError:
            above_floors = 0

        try:
            underground = int(self.underground_floors.get()) if self.underground_floors.get() else 0
        except ValueError:
            underground = 0

        # 現在必要な階のキーセットを作成
        current_floor_keys = set()
        for i in range(above_floors):
            floor_num = above_floors - i
            current_floor_keys.add(f"{floor_num}F")
        for i in range(underground):
            floor_num = i + 1
            current_floor_keys.add(f"B{floor_num}F")

        # 不要な階のデータを削除
        keys_to_remove = [key for key in self.hallway_arch_checks.keys() if key not in current_floor_keys]
        for key in keys_to_remove:
            del self.hallway_arch_checks[key]
            del self.hallway_arch_areas[key]
            del self.hallway_arch_costs[key]
            del self.hallway_arch_processes[key]

        # 新しい階の変数を作成
        for floor_key in current_floor_keys:
            if floor_key not in self.hallway_arch_checks:
                self.hallway_arch_checks[floor_key] = tk.BooleanVar(value=False)
                self.hallway_arch_areas[floor_key] = tk.StringVar(value="0")
                self.hallway_arch_costs[floor_key] = tk.StringVar(value="0")
                self.hallway_arch_processes[floor_key] = tk.StringVar(value="")

        # 建築フレームの内容を更新
        for widget in self.hallway_arch_frame.winfo_children():
            widget.destroy()

        current_row = 0

        # 地上階を上層階から順に表示
        for i in range(above_floors):
            floor_num = above_floors - i
            floor_key = f"{floor_num}F"

            self.add_lounge_item(
                self.hallway_arch_frame,
                current_row,
                f"{floor_num}階通路",
                self.hallway_arch_checks[floor_key],
                self.hallway_arch_areas[floor_key],
                "㎡",
                self.hallway_arch_costs[floor_key]
            )

            # 面積変更時の計算トレース
            self.hallway_arch_areas[floor_key].trace('w', self.update_hallway_arch_costs)
            # チェックボックス変更時の小計更新
            self.hallway_arch_checks[floor_key].trace('w', self.update_hallway_subtotal)

            current_row += 1

        # 地下階を表示
        for i in range(underground):
            floor_num = i + 1
            floor_key = f"B{floor_num}F"

            self.add_lounge_item(
                self.hallway_arch_frame,
                current_row,
                f"地下{floor_num}階通路",
                self.hallway_arch_checks[floor_key],
                self.hallway_arch_areas[floor_key],
                "㎡",
                self.hallway_arch_costs[floor_key]
            )

            # 面積変更時の計算トレース
            self.hallway_arch_areas[floor_key].trace('w', self.update_hallway_arch_costs)
            # チェックボックス変更時の小計更新
            self.hallway_arch_checks[floor_key].trace('w', self.update_hallway_subtotal)

            current_row += 1

        # 建築項目の計算過程フレームを更新
        for widget in self.hallway_arch_process_frame.winfo_children():
            widget.destroy()

        process_text_frame = ttk.Frame(self.hallway_arch_process_frame)
        process_text_frame.pack(fill=tk.BOTH, expand=True)

        # 各階の計算過程ラベルを作成
        sorted_keys = []
        # 地上階を上から順に
        for i in range(above_floors):
            floor_num = above_floors - i
            sorted_keys.append(f"{floor_num}F")
        # 地下階を上から順に
        for i in range(underground):
            floor_num = i + 1
            sorted_keys.append(f"B{floor_num}F")

        for floor_key in sorted_keys:
            if floor_key in self.hallway_arch_processes:
                label = ttk.Label(process_text_frame, textvariable=self.hallway_arch_processes[floor_key],
                                  font=("Courier", 9), justify=tk.LEFT, foreground="blue")
                label.pack(anchor=tk.W, pady=3)

        # 初期計算
        self.update_hallway_arch_costs()
        self.update_hallway_recommended_counts()

    def update_hallway_arch_costs(self, *args):
        """通路タブの建築項目の金額を自動計算"""
        for floor_key in self.hallway_arch_areas.keys():
            try:
                area = float(self.hallway_arch_areas[floor_key].get())
                if area > 0:
                    # 建築初期設定から「通路」の単価を取得
                    height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("通路")

                    if height and floor_cost and ceiling_cost and wall_cost:
                        cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                        self.hallway_arch_costs[floor_key].set(f"{int(cost):,}")

                        # 階数表示用のラベルを生成
                        if floor_key.startswith("B"):
                            floor_label = f"地下{floor_key[1:-1]}階通路"
                        else:
                            floor_label = f"{floor_key[:-1]}階通路"

                        process = (f"【{floor_label} - {grade_name}】\n"
                                   f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                                   f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                                   f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                                   f"       = {int(cost):,}円")
                        self.hallway_arch_processes[floor_key].set(process)
                    else:
                        self.hallway_arch_costs[floor_key].set("0")
                        if floor_key.startswith("B"):
                            floor_label = f"地下{floor_key[1:-1]}階通路"
                        else:
                            floor_label = f"{floor_key[:-1]}階通路"
                        self.hallway_arch_processes[floor_key].set(f"【{floor_label}】単価情報が取得できません")
                else:
                    self.hallway_arch_costs[floor_key].set("0")
                    self.hallway_arch_processes[floor_key].set("")
            except ValueError:
                self.hallway_arch_costs[floor_key].set("0")
                self.hallway_arch_processes[floor_key].set("")

        self.update_hallway_recommended_counts()
        self.update_hallway_subtotal()

    def update_hallway_recommended_counts(self, *args):
        """通路の機械設備・電気設備の推奨個数を計算"""
        try:
            total_area = 0
            for area_var in self.hallway_arch_areas.values():
                try:
                    area = float(area_var.get()) if area_var.get() else 0
                    total_area += area
                except ValueError:
                    pass

            if total_area > 0:
                # 共用部エアコン
                ac_data = self.get_equipment_data_by_abbrev("共用部エアコン")
                if ac_data and ac_data['area_ratio'] > 0:
                    recommended_ac = math.ceil(total_area * ac_data['area_ratio'])
                    self.hallway_mech_ac_recommended.set(f"推奨: {recommended_ac}台")
                else:
                    self.hallway_mech_ac_recommended.set("")

                # スプリンクラー
                sprinkler_data = self.get_equipment_data_by_abbrev("スプリンクラー")
                if sprinkler_data and sprinkler_data['area_ratio'] > 0:
                    recommended_sprinkler = math.ceil(total_area * sprinkler_data['area_ratio'])
                    self.hallway_mech_sprinkler_recommended.set(f"推奨: {recommended_sprinkler}個")
                else:
                    self.hallway_mech_sprinkler_recommended.set("")

                # 消火栓箱
                fire_hose_data = self.get_equipment_data_by_abbrev("消火栓箱")
                if fire_hose_data and fire_hose_data['area_ratio'] > 0:
                    recommended_fire_hose = math.ceil(total_area * fire_hose_data['area_ratio'])
                    self.hallway_mech_fire_hose_recommended.set(f"推奨: {recommended_fire_hose}台")
                else:
                    self.hallway_mech_fire_hose_recommended.set("")

                # LED照明（通路LED）
                led_data = self.get_equipment_data_by_abbrev("通路LED")
                if led_data and led_data['area_ratio'] > 0:
                    recommended_led = math.ceil(total_area * led_data['area_ratio'])
                    self.hallway_elec_led_recommended.set(f"推奨: {recommended_led}個")
                else:
                    self.hallway_elec_led_recommended.set("")

                # 煙感知器
                smoke_data = self.get_equipment_data_by_abbrev("煙感知器")
                if smoke_data and smoke_data['area_ratio'] > 0:
                    recommended_smoke = math.ceil(total_area * smoke_data['area_ratio'])
                    self.hallway_elec_smoke_recommended.set(f"推奨: {recommended_smoke}個")
                else:
                    self.hallway_elec_smoke_recommended.set("")

                # 誘導灯
                exit_light_data = self.get_equipment_data_by_abbrev("誘導灯")
                if exit_light_data and exit_light_data['area_ratio'] > 0:
                    recommended_exit_light = math.ceil(total_area * exit_light_data['area_ratio'])
                    self.hallway_elec_exit_light_recommended.set(f"推奨: {recommended_exit_light}個")
                else:
                    self.hallway_elec_exit_light_recommended.set("")

                # スピーカー
                speaker_data = self.get_equipment_data_by_abbrev("スピーカー")
                if speaker_data and speaker_data['area_ratio'] > 0:
                    recommended_speaker = math.ceil(total_area * speaker_data['area_ratio'])
                    self.hallway_elec_speaker_recommended.set(f"推奨: {recommended_speaker}個")
                else:
                    self.hallway_elec_speaker_recommended.set("")

                # コンセント
                outlet_data = self.get_equipment_data_by_abbrev("コンセント")
                if outlet_data and outlet_data['area_ratio'] > 0:
                    recommended_outlet = math.ceil(total_area * outlet_data['area_ratio'])
                    self.hallway_elec_outlet_recommended.set(f"推奨: {recommended_outlet}個")
                else:
                    self.hallway_elec_outlet_recommended.set("")
            else:
                self.hallway_mech_ac_recommended.set("")
                self.hallway_mech_sprinkler_recommended.set("")
                self.hallway_mech_fire_hose_recommended.set("")
                self.hallway_elec_led_recommended.set("")
                self.hallway_elec_smoke_recommended.set("")
                self.hallway_elec_exit_light_recommended.set("")
                self.hallway_elec_speaker_recommended.set("")
                self.hallway_elec_outlet_recommended.set("")

        except (ValueError, AttributeError):
            pass

    def update_hallway_equipment_costs(self, *args):
        """通路の機械設備・電気設備の金額を計算"""
        mech_process_text = "【機械設備の計算】\n"
        elec_process_text = "【電気設備の計算】\n"

        # 共用部エアコン
        try:
            count = float(self.hallway_mech_ac_count.get()) if self.hallway_mech_ac_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("共用部エアコン")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.hallway_mech_ac_cost.set(f"{int(cost):,}")
                    mech_process_text += f"\n[共用部エアコン]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.hallway_mech_ac_cost.set("0")
        except (ValueError, TypeError):
            self.hallway_mech_ac_cost.set("0")

        # スプリンクラー
        try:
            count = float(self.hallway_mech_sprinkler_count.get()) if self.hallway_mech_sprinkler_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("スプリンクラー")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.hallway_mech_sprinkler_cost.set(f"{int(cost):,}")
                    mech_process_text += f"\n[スプリンクラー]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
                else:
                    self.hallway_mech_sprinkler_cost.set("0")
            else:
                self.hallway_mech_sprinkler_cost.set("0")
        except (ValueError, TypeError):
            self.hallway_mech_sprinkler_cost.set("0")

        # 消火栓箱
        try:
            count = float(self.hallway_mech_fire_hose_count.get()) if self.hallway_mech_fire_hose_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("消火栓箱")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.hallway_mech_fire_hose_cost.set(f"{int(cost):,}")
                    mech_process_text += f"\n[消火栓箱]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
                else:
                    self.hallway_mech_fire_hose_cost.set("0")
            else:
                self.hallway_mech_fire_hose_cost.set("0")
        except (ValueError, TypeError):
            self.hallway_mech_fire_hose_cost.set("0")

        # LED照明
        try:
            count = float(self.hallway_elec_led_count.get()) if self.hallway_elec_led_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("通路LED")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.hallway_elec_led_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[LED照明]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
                else:
                    self.hallway_elec_led_cost.set("0")
            else:
                self.hallway_elec_led_cost.set("0")
        except (ValueError, TypeError):
            self.hallway_elec_led_cost.set("0")

        # 煙感知器
        try:
            count = float(self.hallway_elec_smoke_count.get()) if self.hallway_elec_smoke_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("煙感知器")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.hallway_elec_smoke_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[煙感知器]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
                else:
                    self.hallway_elec_smoke_cost.set("0")
            else:
                self.hallway_elec_smoke_cost.set("0")
        except (ValueError, TypeError):
            self.hallway_elec_smoke_cost.set("0")

        # 誘導灯
        try:
            count = float(self.hallway_elec_exit_light_count.get()) if self.hallway_elec_exit_light_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("誘導灯")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.hallway_elec_exit_light_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[誘導灯]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
                else:
                    self.hallway_elec_exit_light_cost.set("0")
            else:
                self.hallway_elec_exit_light_cost.set("0")
        except (ValueError, TypeError):
            self.hallway_elec_exit_light_cost.set("0")

        # スピーカー
        try:
            count = float(self.hallway_elec_speaker_count.get()) if self.hallway_elec_speaker_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("スピーカー")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.hallway_elec_speaker_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[スピーカー]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
                else:
                    self.hallway_elec_speaker_cost.set("0")
            else:
                self.hallway_elec_speaker_cost.set("0")
        except (ValueError, TypeError):
            self.hallway_elec_speaker_cost.set("0")

        # コンセント
        try:
            count = float(self.hallway_elec_outlet_count.get()) if self.hallway_elec_outlet_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("コンセント")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.hallway_elec_outlet_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[コンセント]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
                else:
                    self.hallway_elec_outlet_cost.set("0")
            else:
                self.hallway_elec_outlet_cost.set("0")
        except (ValueError, TypeError):
            self.hallway_elec_outlet_cost.set("0")

        self.hallway_mech_process.set(mech_process_text)
        self.hallway_elec_process.set(elec_process_text)

        # 小計を更新
        self.update_hallway_subtotal()

    def update_hallway_subtotal(self, *args):
        """通路小計を計算してリアルタイム更新"""
        total = 0

        # 建築項目
        for floor_key in self.hallway_arch_checks.keys():
            if self.hallway_arch_checks[floor_key].get():
                try:
                    cost = float(self.hallway_arch_costs[floor_key].get().replace(',', ''))
                    total += cost
                except (ValueError, AttributeError):
                    pass

        # 機械設備項目
        if self.hallway_mech_ac_check.get():
            try:
                cost = float(self.hallway_mech_ac_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.hallway_mech_sprinkler_check.get():
            try:
                cost = float(self.hallway_mech_sprinkler_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.hallway_mech_fire_hose_check.get():
            try:
                cost = float(self.hallway_mech_fire_hose_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # 電気設備項目
        if self.hallway_elec_led_check.get():
            try:
                cost = float(self.hallway_elec_led_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.hallway_elec_smoke_check.get():
            try:
                cost = float(self.hallway_elec_smoke_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.hallway_elec_exit_light_check.get():
            try:
                cost = float(self.hallway_elec_exit_light_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.hallway_elec_speaker_check.get():
            try:
                cost = float(self.hallway_elec_speaker_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.hallway_elec_outlet_check.get():
            try:
                cost = float(self.hallway_elec_outlet_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # 小計を更新
        self.hallway_subtotal.set(f" {int(total):,}")
    # 通路のタブに関する項目終了

    # アミューズメントのタブに関する項目開始
    def create_amusement_tab(self):
        """アミューズメントタブの作成"""
        canvas = tk.Canvas(self.amusement_frame)
        scrollbar = ttk.Scrollbar(self.amusement_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(main_frame, text="アミューズメント詳細設定", font=title_font).grid(row=0, column=0, columnspan=3,
                                                                                        pady=(0, 20), sticky=tk.W)
        ttk.Label(main_frame, text="面積を入力すると金額が表示されます。計上金額に含める場合は、項目に☑を入れてください。",
                  font=("Arial", 8)).grid(row=1, column=0, columnspan=3, sticky=tk.E, pady=(20, 0))
        row = 2

        # アミューズメント小計
        subtotal_frame = ttk.Frame(main_frame)
        subtotal_frame.grid(row=row, column=0, sticky=tk.W, pady=8)
        ttk.Label(subtotal_frame, text="小計:", font=item_font).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Entry(subtotal_frame, textvariable=self.amusement_subtotal, width=10, state='readonly',
                  font=item_font, justify='right').pack(side=tk.LEFT)
        ttk.Label(subtotal_frame, text=" 円", font=item_font).pack(side=tk.LEFT, padx=(5, 0))
        row += 1

        arch_frame = ttk.LabelFrame(main_frame, text="建築", padding="15")
        arch_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1

        self.add_lounge_item(arch_frame, 0, "卓球コーナー", self.amusement_arch_pingpong_check, self.amusement_arch_pingpong_area, "㎡",
                              self.amusement_arch_pingpong_cost)
        self.add_lounge_item(arch_frame, 1, "キッズスペース", self.amusement_arch_kids_check, self.amusement_arch_kids_area, "㎡",
                              self.amusement_arch_kids_cost)
        self.add_lounge_item(arch_frame, 2, "漫画コーナー", self.amusement_arch_manga_check, self.amusement_arch_manga_area,
                              "㎡", self.amusement_arch_manga_cost)

        process_frame = ttk.LabelFrame(main_frame, text="建築項目の計算過程", padding="10")
        process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        process_text_frame = ttk.Frame(process_frame)
        process_text_frame.pack(fill=tk.BOTH, expand=True)

        pingpong_label = ttk.Label(process_text_frame, textvariable=self.amusement_arch_pingpong_process, font=("Courier", 9),
                                justify=tk.LEFT, foreground="blue")
        pingpong_label.pack(anchor=tk.W, pady=3)

        kids_label = ttk.Label(process_text_frame, textvariable=self.amusement_arch_kids_process, font=("Courier", 9),
                                justify=tk.LEFT, foreground="blue")
        kids_label.pack(anchor=tk.W, pady=3)

        manga_label = ttk.Label(process_text_frame, textvariable=self.amusement_arch_manga_process, font=("Courier", 9),
                                 justify=tk.LEFT, foreground="blue")
        manga_label.pack(anchor=tk.W, pady=3)

        mech_frame = ttk.LabelFrame(main_frame, text="機械設備", padding="15")
        mech_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
        row += 1

        self.add_lounge_item(mech_frame, 0, "共用部エアコン", self.amusement_mech_ac_check, self.amusement_mech_ac_count,
                              "台", self.amusement_mech_ac_cost, self.amusement_mech_ac_recommended)
        self.add_lounge_item(mech_frame, 1, "スプリンクラー", self.amusement_mech_sprinkler_check,
                              self.amusement_mech_sprinkler_count, "個", self.amusement_mech_sprinkler_cost,
                              self.amusement_mech_sprinkler_recommended)
        self.add_lounge_item(mech_frame, 2, "消火栓箱", self.amusement_mech_fire_hose_check,
                              self.amusement_mech_fire_hose_count, "台", self.amusement_mech_fire_hose_cost,
                              self.amusement_mech_fire_hose_recommended)

        mech_process_frame = ttk.LabelFrame(main_frame, text="機械設備の計算過程", padding="10")
        mech_process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        mech_process_label = ttk.Label(mech_process_frame, textvariable=self.amusement_mech_process, font=("Courier", 9),
                                       justify=tk.LEFT, foreground="darkgreen")
        mech_process_label.pack(fill=tk.BOTH, expand=True, anchor=tk.W)

        elec_frame = ttk.LabelFrame(main_frame, text="電気設備", padding="15")
        elec_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
        row += 1

        self.add_lounge_item(elec_frame, 0, "LED照明", self.amusement_elec_led_check, self.amusement_elec_led_count, "個",
                              self.amusement_elec_led_cost, self.amusement_elec_led_recommended)
        self.add_lounge_item(elec_frame, 1, "煙感知器", self.amusement_elec_smoke_check, self.amusement_elec_smoke_count,
                              "個", self.amusement_elec_smoke_cost, self.amusement_elec_smoke_recommended)
        self.add_lounge_item(elec_frame, 2, "誘導灯", self.amusement_elec_exit_light_check,
                              self.amusement_elec_exit_light_count, "個", self.amusement_elec_exit_light_cost,
                              self.amusement_elec_exit_light_recommended)
        self.add_lounge_item(elec_frame, 3, "スピーカー", self.amusement_elec_speaker_check,
                              self.amusement_elec_speaker_count, "個", self.amusement_elec_speaker_cost,
                              self.amusement_elec_speaker_recommended)
        self.add_lounge_item(elec_frame, 4, "コンセント", self.amusement_elec_outlet_check, self.amusement_elec_outlet_count,
                              "個", self.amusement_elec_outlet_cost, self.amusement_elec_outlet_recommended)

        elec_process_frame = ttk.LabelFrame(main_frame, text="電気設備の計算過程", padding="10")
        elec_process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        elec_process_label = ttk.Label(elec_process_frame, textvariable=self.amusement_elec_process, font=("Courier", 9),
                                       justify=tk.LEFT, foreground="darkgreen")
        elec_process_label.pack(fill=tk.BOTH, expand=True, anchor=tk.W)

    def update_amusement_subtotal(self, *args):
        """アミューズメント小計を計算してリアルタイム更新"""
        total = 0

        # 建築項目
        if self.amusement_arch_pingpong_check.get():
            try:
                cost = float(self.amusement_arch_pingpong_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.amusement_arch_kids_check.get():
            try:
                cost = float(self.amusement_arch_kids_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.amusement_arch_manga_check.get():
            try:
                cost = float(self.amusement_arch_manga_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # 機械設備項目
        if self.amusement_mech_ac_check.get():
            try:
                cost = float(self.amusement_mech_ac_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.amusement_mech_sprinkler_check.get():
            try:
                cost = float(self.amusement_mech_sprinkler_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.amusement_mech_fire_hose_check.get():
            try:
                cost = float(self.amusement_mech_fire_hose_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # 電気設備項目
        if self.amusement_elec_led_check.get():
            try:
                cost = float(self.amusement_elec_led_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.amusement_elec_smoke_check.get():
            try:
                cost = float(self.amusement_elec_smoke_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.amusement_elec_exit_light_check.get():
            try:
                cost = float(self.amusement_elec_exit_light_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.amusement_elec_speaker_check.get():
            try:
                cost = float(self.amusement_elec_speaker_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.amusement_elec_outlet_check.get():
            try:
                cost = float(self.amusement_elec_outlet_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # 小計を更新
        self.amusement_subtotal.set(f" {int(total):,}")

    def update_amusement_arch_costs(self, *args):
        """アミューズメントタブの建築項目の金額を自動計算"""
        try:
            area = float(self.amusement_arch_pingpong_area.get())
            if area > 0:
                # 初期設定から室用途を照合する検索キーワード
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("卓球コーナー")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    self.amusement_arch_pingpong_cost.set(f"{int(cost):,}")

                    process = (f"【卓球コーナー - {grade_name}】\n"
                               f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                               f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                               f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                               f"       = {int(cost):,}円")
                    self.amusement_arch_pingpong_process.set(process)
                else:
                    self.amusement_arch_pingpong_cost.set("0")
                    self.amusement_arch_pingpong_process.set("【卓球コーナー】単価情報が取得できません")
            else:
                self.amusement_arch_pingpong_cost.set("0")
                self.amusement_arch_pingpong_process.set("")
        except ValueError:
            self.amusement_arch_pingpong_cost.set("0")
            self.amusement_arch_pingpong_process.set("")

        try:
            area = float(self.amusement_arch_kids_area.get())
            if area > 0:
                #初期設定から室用途を照合する検索キーワード
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("キッズスペース")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    self.amusement_arch_kids_cost.set(f"{int(cost):,}")

                    process = (f"【キッズスペース - {grade_name}】\n"
                               f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                               f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                               f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                               f"       = {int(cost):,}円")
                    self.amusement_arch_kids_process.set(process)
                else:
                    self.amusement_arch_kids_cost.set("0")
                    self.amusement_arch_kids_process.set("【キッズスペース】単価情報が取得できません")
            else:
                self.amusement_arch_kids_cost.set("0")
                self.amusement_arch_kids_process.set("")
        except ValueError:
            self.amusement_arch_kids_cost.set("0")
            self.amusement_arch_kids_process.set("")

        try:
            area = float(self.amusement_arch_manga_area.get())
            if area > 0:
                # 初期設定から室用途を照合する検索キーワード
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("漫画コーナー")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    self.amusement_arch_manga_cost.set(f"{int(cost):,}")

                    process = (f"【漫画コーナー - {grade_name}】\n"
                               f"  金額 = {area}×{floor_cost:,} + {area}×{ceiling_cost:,} + "
                               f"{area:.2f}^0.5×4×{height}×{wall_cost:,}\n"
                               f"       = {area * floor_cost:,.0f} + {area * ceiling_cost:,.0f} + {(area ** 0.5) * 4 * height * wall_cost:,.0f}\n"
                               f"       = {int(cost):,}円")
                    self.amusement_arch_manga_process.set(process)
                else:
                    self.amusement_arch_manga_cost.set("0")
                    self.amusement_arch_manga_process.set("【漫画コーナー】単価情報が取得できません")
            else:
                self.amusement_arch_manga_cost.set("0")
                self.amusement_arch_manga_process.set("")
        except ValueError:
            self.amusement_arch_manga_cost.set("0")
            self.amusement_arch_manga_process.set("")

        self.update_amusement_recommended_counts()
        self.update_amusement_subtotal()

    def update_amusement_recommended_counts(self, *args):
        """アミューズメントの機械設備・電気設備の推奨個数を計算"""
        try:
            pingpong_area = float(self.amusement_arch_pingpong_area.get()) if self.amusement_arch_pingpong_area.get() else 0
            kids_area = float(self.amusement_arch_kids_area.get()) if self.amusement_arch_kids_area.get() else 0
            manga_area = float(self.amusement_arch_manga_area.get()) if self.amusement_arch_manga_area.get() else 0
            total_area = pingpong_area + kids_area + manga_area

            if total_area > 0:
                ac_data = self.get_equipment_data_by_abbrev("共用部エアコン")
                if ac_data and ac_data['area_ratio'] > 0:
                    recommended_ac = math.ceil(total_area * ac_data['area_ratio'])
                    self.amusement_mech_ac_recommended.set(f"推奨: {recommended_ac}台")
                else:
                    self.amusement_mech_ac_recommended.set("")

                sprinkler_data = self.get_equipment_data_by_abbrev("スプリンクラー")
                if sprinkler_data and sprinkler_data['area_ratio'] > 0:
                    recommended_sprinkler = math.ceil(total_area * sprinkler_data['area_ratio'])
                    self.amusement_mech_sprinkler_recommended.set(f"推奨: {recommended_sprinkler}個")
                else:
                    self.amusement_mech_sprinkler_recommended.set("")

                fire_hose_data = self.get_equipment_data_by_abbrev("消火栓箱")
                if fire_hose_data and fire_hose_data['area_ratio'] > 0:
                    recommended_fire_hose = math.ceil(total_area * fire_hose_data['area_ratio'])
                    self.amusement_mech_fire_hose_recommended.set(f"推奨: {recommended_fire_hose}台")
                else:
                    self.amusement_mech_fire_hose_recommended.set("")

                led_found = False
                for abbrev in ["エントランスLED", "通路LED", "レストランLED"]:
                    led_data = self.get_equipment_data_by_abbrev(abbrev)
                    if led_data and led_data['area_ratio'] > 0:
                        recommended_led = math.ceil(total_area * led_data['area_ratio'])
                        self.amusement_elec_led_recommended.set(f"推奨: {recommended_led}個")
                        led_found = True
                        break
                if not led_found:
                    self.amusement_elec_led_recommended.set("")

                smoke_data = self.get_equipment_data_by_abbrev("煙感知器")
                if smoke_data and smoke_data['area_ratio'] > 0:
                    recommended_smoke = math.ceil(total_area * smoke_data['area_ratio'])
                    self.amusement_elec_smoke_recommended.set(f"推奨: {recommended_smoke}個")
                else:
                    self.amusement_elec_smoke_recommended.set("")

                exit_light_data = self.get_equipment_data_by_abbrev("誘導灯")
                if exit_light_data and exit_light_data['area_ratio'] > 0:
                    recommended_exit_light = math.ceil(total_area * exit_light_data['area_ratio'])
                    self.amusement_elec_exit_light_recommended.set(f"推奨: {recommended_exit_light}個")
                else:
                    self.amusement_elec_exit_light_recommended.set("")

                speaker_data = self.get_equipment_data_by_abbrev("スピーカー")
                if speaker_data and speaker_data['area_ratio'] > 0:
                    recommended_speaker = math.ceil(total_area * speaker_data['area_ratio'])
                    self.amusement_elec_speaker_recommended.set(f"推奨: {recommended_speaker}個")
                else:
                    self.amusement_elec_speaker_recommended.set("")

                outlet_data = self.get_equipment_data_by_abbrev("コンセント")
                if outlet_data and outlet_data['area_ratio'] > 0:
                    recommended_outlet = math.ceil(total_area * outlet_data['area_ratio'])
                    self.amusement_elec_outlet_recommended.set(f"推奨: {recommended_outlet}個")
                else:
                    self.amusement_elec_outlet_recommended.set("")
            else:
                self.amusement_mech_ac_recommended.set("")
                self.amusement_mech_sprinkler_recommended.set("")
                self.amusement_mech_fire_hose_recommended.set("")
                self.amusement_elec_led_recommended.set("")
                self.amusement_elec_smoke_recommended.set("")
                self.amusement_elec_exit_light_recommended.set("")
                self.amusement_elec_speaker_recommended.set("")
                self.amusement_elec_outlet_recommended.set("")

        except (ValueError, AttributeError):
            pass

    def update_amusement_equipment_costs(self, *args):
        """アミューズメントの機械設備・電気設備の金額を計算"""
        mech_process_text = "【機械設備の計算】\n"
        elec_process_text = "【電気設備の計算】\n"

        try:
            count = float(self.amusement_mech_ac_count.get()) if self.amusement_mech_ac_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("共用部エアコン")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.amusement_mech_ac_cost.set(f"{int(cost):,}")
                    mech_process_text += f"\n[共用部エアコン]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.amusement_mech_ac_cost.set("0")
        except (ValueError, TypeError):
            self.amusement_mech_ac_cost.set("0")

        try:
            count = float(self.amusement_mech_sprinkler_count.get()) if self.amusement_mech_sprinkler_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("スプリンクラー")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.amusement_mech_sprinkler_cost.set(f"{int(cost):,}")
                    mech_process_text += f"\n[スプリンクラー]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.amusement_mech_sprinkler_cost.set("0")
        except (ValueError, TypeError):
            self.amusement_mech_sprinkler_cost.set("0")

        try:
            count = float(self.amusement_mech_fire_hose_count.get()) if self.amusement_mech_fire_hose_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("消火栓箱")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.amusement_mech_fire_hose_cost.set(f"{int(cost):,}")
                    mech_process_text += f"\n[消火栓箱]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.amusement_mech_fire_hose_cost.set("0")
        except (ValueError, TypeError):
            self.amusement_mech_fire_hose_cost.set("0")

        try:
            count = float(self.amusement_elec_led_count.get()) if self.amusement_elec_led_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("エントランスLED")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.amusement_elec_led_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[LED照明]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.amusement_elec_led_cost.set("0")
        except (ValueError, TypeError):
            self.amusement_elec_led_cost.set("0")

        try:
            count = float(self.amusement_elec_smoke_count.get()) if self.amusement_elec_smoke_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("煙感知器")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.amusement_elec_smoke_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[煙感知器]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.amusement_elec_smoke_cost.set("0")
        except (ValueError, TypeError):
            self.amusement_elec_smoke_cost.set("0")

        try:
            count = float(self.amusement_elec_exit_light_count.get()) if self.amusement_elec_exit_light_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("誘導灯")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.amusement_elec_exit_light_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[誘導灯]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.amusement_elec_exit_light_cost.set("0")
        except (ValueError, TypeError):
            self.amusement_elec_exit_light_cost.set("0")

        try:
            count = float(self.amusement_elec_speaker_count.get()) if self.amusement_elec_speaker_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("スピーカー")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.amusement_elec_speaker_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[スピーカー]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.amusement_elec_speaker_cost.set("0")
        except (ValueError, TypeError):
            self.amusement_elec_speaker_cost.set("0")

        try:
            count = float(self.amusement_elec_outlet_count.get()) if self.amusement_elec_outlet_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("コンセント")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.amusement_elec_outlet_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[コンセント]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.amusement_elec_outlet_cost.set("0")
        except (ValueError, TypeError):
            self.amusement_elec_outlet_cost.set("0")

        self.amusement_mech_process.set(mech_process_text)
        self.amusement_elec_process.set(elec_process_text)

        # 小計を更新
        self.update_amusement_subtotal()
    # アミューズメントのタブに関する項目終了

    # 客室のタブに関する項目開始
    def create_guest_room_tab(self):
        """客室タブの作成"""
        canvas = tk.Canvas(self.guest_room_frame)
        scrollbar = ttk.Scrollbar(self.guest_room_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(main_frame, text="客室詳細設定", font=title_font).grid(row=0, column=0, columnspan=3,
                                                                                    pady=(0, 20), sticky=tk.W)
        ttk.Label(main_frame, text="面積を入力すると金額が表示されます。計上金額に含める場合は、項目に☑を入れてください。",
                  font=("Arial", 8)).grid(row=1, column=0, columnspan=3, sticky=tk.E, pady=(20, 0))
        row = 2

        # 客室小計
        subtotal_frame = ttk.Frame(main_frame)
        subtotal_frame.grid(row=row, column=0, sticky=tk.W, pady=8)
        ttk.Label(subtotal_frame, text="小計:", font=item_font).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Entry(subtotal_frame, textvariable=self.guest_subtotal, width=10, state='readonly',
                  font=item_font, justify='right').pack(side=tk.LEFT)
        ttk.Label(subtotal_frame, text=" 円", font=item_font).pack(side=tk.LEFT, padx=(5, 0))
        row += 1

        # 建築項目
        arch_frame = ttk.LabelFrame(main_frame, text="建築", padding="15")
        arch_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        row += 1

        # 和室(10-15J)の行 - 部屋数表示を追加
        arch_row = 0
        ttk.Checkbutton(arch_frame, text="和室(10-15J)", variable=self.guest_arch_japanese_check).grid(
            row=arch_row, column=0, sticky=tk.W, padx=(0, 10), pady=3)
        ttk.Entry(arch_frame, textvariable=self.guest_arch_japanese_area, width=10, justify=tk.RIGHT).grid(
            row=arch_row, column=1, sticky=tk.W, padx=(5, 5), pady=3)
        ttk.Label(arch_frame, text="㎡").grid(row=arch_row, column=2, sticky=tk.W, padx=(5, 10), pady=3)

        # 部屋数表示ラベル（動的更新）
        self.guest_japanese_room_count_label = ttk.Label(arch_frame, text="(0室)",
                                                         font=("Arial", 9), foreground="blue")
        self.guest_japanese_room_count_label.grid(row=arch_row, column=3, sticky=tk.W, padx=(5, 10), pady=3)

        ttk.Entry(arch_frame, textvariable=self.guest_arch_japanese_cost, width=12, state='readonly',
                  justify=tk.RIGHT).grid(row=arch_row, column=4, sticky=tk.W, padx=(5, 5), pady=3)
        ttk.Label(arch_frame, text="円").grid(row=arch_row, column=5, sticky=tk.W, pady=3)

        # 和室・洋室の行 - 部屋数表示を追加
        arch_row = 1
        ttk.Checkbutton(arch_frame, text="和室・洋室", variable=self.guest_arch_japanese_western_check).grid(
            row=arch_row, column=0, sticky=tk.W, padx=(0, 10), pady=3)
        ttk.Entry(arch_frame, textvariable=self.guest_arch_japanese_western_area, width=10, justify=tk.RIGHT).grid(
            row=arch_row, column=1, sticky=tk.W, padx=(5, 5), pady=3)
        ttk.Label(arch_frame, text="㎡").grid(row=arch_row, column=2, sticky=tk.W, padx=(5, 10), pady=3)

        # 部屋数表示ラベル（動的更新）
        self.guest_japanese_western_room_count_label = ttk.Label(arch_frame, text="(0室)",
                                                                 font=("Arial", 9), foreground="blue")
        self.guest_japanese_western_room_count_label.grid(row=arch_row, column=3, sticky=tk.W, padx=(5, 10), pady=3)

        ttk.Entry(arch_frame, textvariable=self.guest_arch_japanese_western_cost, width=12, state='readonly',
                  justify=tk.RIGHT).grid(row=arch_row, column=4, sticky=tk.W, padx=(5, 5), pady=3)
        ttk.Label(arch_frame, text="円").grid(row=arch_row, column=5, sticky=tk.W, pady=3)

        # 和ベッドの行 - 部屋数表示を追加
        arch_row = 2
        ttk.Checkbutton(arch_frame, text="和ベッド", variable=self.guest_arch_japanese_bed_check).grid(
            row=arch_row, column=0, sticky=tk.W, padx=(0, 10), pady=3)
        ttk.Entry(arch_frame, textvariable=self.guest_arch_japanese_bed_area, width=10, justify=tk.RIGHT).grid(
            row=arch_row, column=1, sticky=tk.W, padx=(5, 5), pady=3)
        ttk.Label(arch_frame, text="㎡").grid(row=arch_row, column=2, sticky=tk.W, padx=(5, 10), pady=3)

        # 部屋数表示ラベル（動的更新）
        self.guest_japanese_bed_room_count_label = ttk.Label(arch_frame, text="(0室)",
                                                             font=("Arial", 9), foreground="blue")
        self.guest_japanese_bed_room_count_label.grid(row=arch_row, column=3, sticky=tk.W, padx=(5, 10), pady=3)

        ttk.Entry(arch_frame, textvariable=self.guest_arch_japanese_bed_cost, width=12, state='readonly',
                  justify=tk.RIGHT).grid(row=arch_row, column=4, sticky=tk.W, padx=(5, 5), pady=3)
        ttk.Label(arch_frame, text="円").grid(row=arch_row, column=5, sticky=tk.W, pady=3)

        # 洋室の行 - 部屋数表示を追加
        arch_row = 3
        ttk.Checkbutton(arch_frame, text="洋室", variable=self.guest_arch_western_check).grid(
            row=arch_row, column=0, sticky=tk.W, padx=(0, 10), pady=3)
        ttk.Entry(arch_frame, textvariable=self.guest_arch_western_area, width=10, justify=tk.RIGHT).grid(
            row=arch_row, column=1, sticky=tk.W, padx=(5, 5), pady=3)
        ttk.Label(arch_frame, text="㎡").grid(row=arch_row, column=2, sticky=tk.W, padx=(5, 10), pady=3)

        # 部屋数表示ラベル（動的更新）
        self.guest_western_room_count_label = ttk.Label(arch_frame, text="(0室)",
                                                        font=("Arial", 9), foreground="blue")
        self.guest_western_room_count_label.grid(row=arch_row, column=3, sticky=tk.W, padx=(5, 10), pady=3)

        ttk.Entry(arch_frame, textvariable=self.guest_arch_western_cost, width=12, state='readonly',
                  justify=tk.RIGHT).grid(row=arch_row, column=4, sticky=tk.W, padx=(5, 5), pady=3)
        ttk.Label(arch_frame, text="円").grid(row=arch_row, column=5, sticky=tk.W, pady=3)

        # 建築項目の計算過程
        process_frame = ttk.LabelFrame(main_frame, text="建築項目の計算過程", padding="10")
        process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        process_text_frame = ttk.Frame(process_frame)
        process_text_frame.pack(fill=tk.BOTH, expand=True)

        japanese_label = ttk.Label(process_text_frame, textvariable=self.guest_arch_japanese_process,
                                   font=("Courier", 9), justify=tk.LEFT, foreground="blue")
        japanese_label.pack(anchor=tk.W, pady=3)

        japanese_western_label = ttk.Label(process_text_frame, textvariable=self.guest_arch_japanese_western_process,
                                           font=("Courier", 9), justify=tk.LEFT, foreground="blue")
        japanese_western_label.pack(anchor=tk.W, pady=3)

        japanese_bed_label = ttk.Label(process_text_frame, textvariable=self.guest_arch_japanese_bed_process,
                                       font=("Courier", 9), justify=tk.LEFT, foreground="blue")
        japanese_bed_label.pack(anchor=tk.W, pady=3)

        western_label = ttk.Label(process_text_frame, textvariable=self.guest_arch_western_process,
                                  font=("Courier", 9), justify=tk.LEFT, foreground="blue")
        western_label.pack(anchor=tk.W, pady=3)

        # 機械設備
        mech_frame = ttk.LabelFrame(main_frame, text="機械設備", padding="15")
        mech_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
        row += 1

        self.add_lounge_item(mech_frame, 0, "個別エアコン", self.guest_mech_ac_check, self.guest_mech_ac_count,
                              "台", self.guest_mech_ac_cost, self.guest_mech_ac_recommended)
        self.add_lounge_item(mech_frame, 1, "洗面器", self.guest_mech_wash_basin_check,
                              self.guest_mech_wash_basin_count, "個", self.guest_mech_wash_basin_cost,
                              self.guest_mech_wash_basin_recommended)
        self.add_lounge_item(mech_frame, 2, "スプリンクラー", self.guest_mech_sprinkler_check,
                              self.guest_mech_sprinkler_count, "個", self.guest_mech_sprinkler_cost,
                              self.guest_mech_sprinkler_recommended)

        # 機械設備の計算過程
        mech_process_frame = ttk.LabelFrame(main_frame, text="機械設備の計算過程", padding="10")
        mech_process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        mech_process_label = ttk.Label(mech_process_frame, textvariable=self.guest_mech_process, font=("Courier", 9),
                                       justify=tk.LEFT, foreground="darkgreen")
        mech_process_label.pack(fill=tk.BOTH, expand=True, anchor=tk.W)

        # 電気設備
        elec_frame = ttk.LabelFrame(main_frame, text="電気設備", padding="15")
        elec_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
        row += 1

        self.add_lounge_item(elec_frame, 0, "主室照明", self.guest_elec_main_light_check,
                              self.guest_elec_main_light_count, "個",
                              self.guest_elec_main_light_cost, self.guest_elec_main_light_recommended)
        self.add_lounge_item(elec_frame, 1, "煙感知器", self.guest_elec_smoke_check, self.guest_elec_smoke_count,
                              "個", self.guest_elec_smoke_cost, self.guest_elec_smoke_recommended)
        self.add_lounge_item(elec_frame, 2, "熱感知器", self.guest_elec_heat_detector_check,
                              self.guest_elec_heat_detector_count, "個", self.guest_elec_heat_detector_cost,
                              self.guest_elec_heat_detector_recommended)

        self.add_lounge_item(elec_frame, 3, "スピーカー", self.guest_elec_speaker_check,
                              self.guest_elec_speaker_count, "個", self.guest_elec_speaker_cost,
                              self.guest_elec_speaker_recommended)
        self.add_lounge_item(elec_frame, 4, "コンセント", self.guest_elec_outlet_check, self.guest_elec_outlet_count,
                              "個", self.guest_elec_outlet_cost, self.guest_elec_outlet_recommended)

        # 電気設備の計算過程
        elec_process_frame = ttk.LabelFrame(main_frame, text="電気設備の計算過程", padding="10")
        elec_process_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 15))
        row += 1

        elec_process_label = ttk.Label(elec_process_frame, textvariable=self.guest_elec_process, font=("Courier", 9),
                                       justify=tk.LEFT, foreground="darkgreen")
        elec_process_label.pack(fill=tk.BOTH, expand=True, anchor=tk.W)

        # 以下、既存の客室タイプ別構成など
        row += 1
        ttk.Label(main_frame, text="客室タイプ別構成", font=item_font).grid(row=row, column=0, columnspan=3,
                                                                                        pady=(20, 5), sticky=tk.W)
        row += 1

        room_types = [
            ("和室(10-15畳):", self.capacity_japanese_room),
            ("和室・洋室:", self.capacity_japanese_western_room),
            ("和ベッド:", self.capacity_japanese_bed_room),
            ("洋室:", self.capacity_western_room),
        ]

        for label_text, var in room_types:
            ttk.Label(main_frame, text=label_text).grid(row=row, column=0, sticky=tk.W, pady=5)
            ttk.Entry(main_frame, textvariable=var, width=5).grid(row=row, column=1, sticky=tk.W, pady=5)
            ttk.Label(main_frame, text="人/室").grid(row=row, column=2, sticky=tk.W, pady=5)
            row += 1

        ttk.Label(main_frame, text="共通客室設備", font=item_font).grid(row=row, column=0, columnspan=3,
                                                                                    pady=(20, 5), sticky=tk.W)
        row += 1

        equipment_list = [
            ("TV棚", self.guest_room_tv_cabinet),
            ("ヘッドボード", self.guest_room_headboard),
            ("露付客室", self.guest_private_open_bath),
        ]

        for label_text, var in equipment_list:
            ttk.Checkbutton(main_frame, text=label_text, variable=var).grid(row=row, column=0, columnspan=3,
                                                                            sticky=tk.W, pady=3)
            row += 1

        self.capacity_japanese_room.trace('w', self.update_total_guests_realtime)
        self.capacity_japanese_western_room.trace('w', self.update_total_guests_realtime)
        self.capacity_japanese_bed_room.trace('w', self.update_total_guests_realtime)
        self.capacity_western_room.trace('w', self.update_total_guests_realtime)

    def get_selected_spec_grade_index(self):
        """コスト計算タブで選択されている仕様グレードのインデックスを取得"""
        try:
            selection = self.spec_listbox.curselection()
            if selection:
                return selection[0]
            return 0
        except:
            return 0

    def get_equipment_data_by_abbrev(self, abbrev):
        """略称で設備機器データを検索"""
        for data_vars in self.equipment_unit_data:
            no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var, new_equipment_cost_var, area_ratio_var = data_vars
            if abbrev_var.get() == abbrev:
                try:
                    return {
                        'install_hours': float(install_hours_var.get()) if install_hours_var.get() else 0,
                        'labor_cost': float(labor_cost_var.get()) if labor_cost_var.get() else 0,
                        'misc_material': float(misc_material_var.get()) if misc_material_var.get() else 0,
                        'new_equipment_cost': float(new_equipment_cost_var.get()) if new_equipment_cost_var.get() else 0,
                        'area_ratio': float(area_ratio_var.get()) if area_ratio_var.get() else 0
                    }
                except (ValueError, AttributeError):
                    return None
        return None

    def update_guest_recommended_counts(self, *args):
        """客室の機械設備・電気設備の推奨個数を計算"""
        try:
            # 各部屋タイプの面積を取得
            japanese_area = float(self.guest_arch_japanese_area.get()) if self.guest_arch_japanese_area.get() else 0
            japanese_western_area = float(
                self.guest_arch_japanese_western_area.get()) if self.guest_arch_japanese_western_area.get() else 0
            japanese_bed_area = float(
                self.guest_arch_japanese_bed_area.get()) if self.guest_arch_japanese_bed_area.get() else 0
            western_area = float(self.guest_arch_western_area.get()) if self.guest_arch_western_area.get() else 0

            # 各部屋タイプの客室数を取得
            japanese_count = int(self.japanese_room_count.get()) if self.japanese_room_count.get() else 0
            japanese_western_count = int(
                self.japanese_western_room_count.get()) if self.japanese_western_room_count.get() else 0
            japanese_bed_count = int(self.japanese_bed_room_count.get()) if self.japanese_bed_room_count.get() else 0
            western_count = int(self.western_room_count.get()) if self.western_room_count.get() else 0

            # 総面積 = 各部屋タイプの面積 × 客室数の合計
            total_area = (japanese_area * japanese_count +
                          japanese_western_area * japanese_western_count +
                          japanese_bed_area * japanese_bed_count +
                          western_area * western_count)

            if total_area > 0:
                # 個別エアコン（旧：共用部エアコン）
                ac_data = self.get_equipment_data_by_abbrev("個別エアコン")
                if ac_data and ac_data['area_ratio'] > 0:
                    recommended_ac = math.ceil(total_area * ac_data['area_ratio'])
                    self.guest_mech_ac_recommended.set(f"推奨: {recommended_ac}台")
                else:
                    self.guest_mech_ac_recommended.set("")

                # 洗面器（旧：スプリンクラー）
                wash_basin_data = self.get_equipment_data_by_abbrev("洗面器")
                if wash_basin_data and wash_basin_data['area_ratio'] > 0:
                    recommended_wash_basin = math.ceil(total_area * wash_basin_data['area_ratio'])
                    self.guest_mech_wash_basin_recommended.set(f"推奨: {recommended_wash_basin}個")
                else:
                    self.guest_mech_wash_basin_recommended.set("")

                # スプリンクラー（旧：消火栓箱）
                sprinkler_data = self.get_equipment_data_by_abbrev("スプリンクラー")
                if sprinkler_data and sprinkler_data['area_ratio'] > 0:
                    recommended_sprinkler = math.ceil(total_area * sprinkler_data['area_ratio'])
                    self.guest_mech_sprinkler_recommended.set(f"推奨: {recommended_sprinkler}個")
                else:
                    self.guest_mech_sprinkler_recommended.set("")

                # 主室照明（旧：LED照明）
                main_light_data = self.get_equipment_data_by_abbrev("主室照明")
                if main_light_data and main_light_data['area_ratio'] > 0:
                    recommended_main_light = math.ceil(total_area * main_light_data['area_ratio'])
                    self.guest_elec_main_light_recommended.set(f"推奨: {recommended_main_light}個")
                else:
                    self.guest_elec_main_light_recommended.set("")

                # 煙感知器
                smoke_data = self.get_equipment_data_by_abbrev("煙感知器")
                if smoke_data and smoke_data['area_ratio'] > 0:
                    recommended_smoke = math.ceil(total_area * smoke_data['area_ratio'])
                    self.guest_elec_smoke_recommended.set(f"推奨: {recommended_smoke}個")
                else:
                    self.guest_elec_smoke_recommended.set("")

                # 熱感知器（旧：誘導灯）
                heat_detector_data = self.get_equipment_data_by_abbrev("熱感知器")
                if heat_detector_data and heat_detector_data['area_ratio'] > 0:
                    recommended_heat_detector = math.ceil(total_area * heat_detector_data['area_ratio'])
                    self.guest_elec_heat_detector_recommended.set(f"推奨: {recommended_heat_detector}個")
                else:
                    self.guest_elec_heat_detector_recommended.set("")

                # スピーカー
                speaker_data = self.get_equipment_data_by_abbrev("スピーカー")
                if speaker_data and speaker_data['area_ratio'] > 0:
                    recommended_speaker = math.ceil(total_area * speaker_data['area_ratio'])
                    self.guest_elec_speaker_recommended.set(f"推奨: {recommended_speaker}個")
                else:
                    self.guest_elec_speaker_recommended.set("")

                # コンセント
                outlet_data = self.get_equipment_data_by_abbrev("コンセント")
                if outlet_data and outlet_data['area_ratio'] > 0:
                    recommended_outlet = math.ceil(total_area * outlet_data['area_ratio'])
                    self.guest_elec_outlet_recommended.set(f"推奨: {recommended_outlet}個")
                else:
                    self.guest_elec_outlet_recommended.set("")
            else:
                self.guest_mech_ac_recommended.set("")
                self.guest_mech_wash_basin_recommended.set("")
                self.guest_mech_sprinkler_recommended.set("")
                self.guest_elec_main_light_recommended.set("")
                self.guest_elec_smoke_recommended.set("")
                self.guest_elec_heat_detector_recommended.set("")
                self.guest_elec_speaker_recommended.set("")
                self.guest_elec_outlet_recommended.set("")

        except (ValueError, AttributeError):
            pass

    def update_guest_equipment_costs(self, *args):
        """客室の機械設備・電気設備の金額を計算"""
        mech_process_text = "【機械設備の計算】\n"
        elec_process_text = "【電気設備の計算】\n"

        # 個別エアコン
        try:
            count = float(self.guest_mech_ac_count.get()) if self.guest_mech_ac_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("個別エアコン")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.guest_mech_ac_cost.set(f"{int(cost):,}")
                    mech_process_text += f"\n[個別エアコン]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.guest_mech_ac_cost.set("0")
        except (ValueError, TypeError):
            self.guest_mech_ac_cost.set("0")

        # 洗面器
        try:
            count = float(self.guest_mech_wash_basin_count.get()) if self.guest_mech_wash_basin_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("洗面器")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.guest_mech_wash_basin_cost.set(f"{int(cost):,}")
                    mech_process_text += f"\n[洗面器]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.guest_mech_wash_basin_cost.set("0")
        except (ValueError, TypeError):
            self.guest_mech_wash_basin_cost.set("0")

        # スプリンクラー
        try:
            count = float(self.guest_mech_sprinkler_count.get()) if self.guest_mech_sprinkler_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("スプリンクラー")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.guest_mech_sprinkler_cost.set(f"{int(cost):,}")
                    mech_process_text += f"\n[スプリンクラー]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.guest_mech_sprinkler_cost.set("0")
        except (ValueError, TypeError):
            self.guest_mech_sprinkler_cost.set("0")

        # 主室照明
        try:
            count = float(self.guest_elec_main_light_count.get()) if self.guest_elec_main_light_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("主室照明")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.guest_elec_main_light_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[主室照明]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.guest_elec_main_light_cost.set("0")
        except (ValueError, TypeError):
            self.guest_elec_main_light_cost.set("0")

        # 煙感知器
        try:
            count = float(self.guest_elec_smoke_count.get()) if self.guest_elec_smoke_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("煙感知器")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.guest_elec_smoke_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[煙感知器]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.guest_elec_smoke_cost.set("0")
        except (ValueError, TypeError):
            self.guest_elec_smoke_cost.set("0")

        # 熱感知器
        try:
            count = float(self.guest_elec_heat_detector_count.get()) if self.guest_elec_heat_detector_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("熱感知器")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.guest_elec_heat_detector_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[熱感知器]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.guest_elec_heat_detector_cost.set("0")
        except (ValueError, TypeError):
            self.guest_elec_heat_detector_cost.set("0")

        # スピーカー
        try:
            count = float(self.guest_elec_speaker_count.get()) if self.guest_elec_speaker_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("スピーカー")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.guest_elec_speaker_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[スピーカー]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.guest_elec_speaker_cost.set("0")
        except (ValueError, TypeError):
            self.guest_elec_speaker_cost.set("0")

        # コンセント
        try:
            count = float(self.guest_elec_outlet_count.get()) if self.guest_elec_outlet_count.get() else 0
            if count > 0:
                data = self.get_equipment_data_by_abbrev("コンセント")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.guest_elec_outlet_cost.set(f"{int(cost):,}")
                    elec_process_text += f"\n[コンセント]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.guest_elec_outlet_cost.set("0")
        except (ValueError, TypeError):
            self.guest_elec_outlet_cost.set("0")

        self.guest_mech_process.set(mech_process_text)
        self.guest_elec_process.set(elec_process_text)

        # 小計を更新
        self.update_guest_subtotal()

    def update_guest_subtotal(self, *args):
        """客室小計を計算してリアルタイム更新"""
        total = 0

        # 建築項目
        if self.guest_arch_japanese_check.get():
            try:
                cost = float(self.guest_arch_japanese_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.guest_arch_japanese_western_check.get():
            try:
                cost = float(self.guest_arch_japanese_western_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.guest_arch_japanese_bed_check.get():
            try:
                cost = float(self.guest_arch_japanese_bed_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.guest_arch_western_check.get():
            try:
                cost = float(self.guest_arch_western_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # 機械設備項目
        if self.guest_mech_ac_check.get():
            try:
                cost = float(self.guest_mech_ac_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.guest_mech_wash_basin_check.get():
            try:
                cost = float(self.guest_mech_wash_basin_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.guest_mech_sprinkler_check.get():
            try:
                cost = float(self.guest_mech_sprinkler_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # 電気設備項目
        if self.guest_elec_main_light_check.get():
            try:
                cost = float(self.guest_elec_main_light_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.guest_elec_smoke_check.get():
            try:
                cost = float(self.guest_elec_smoke_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.guest_elec_heat_detector_check.get():
            try:
                cost = float(self.guest_elec_heat_detector_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.guest_elec_speaker_check.get():
            try:
                cost = float(self.guest_elec_speaker_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        if self.guest_elec_outlet_check.get():
            try:
                cost = float(self.guest_elec_outlet_cost.get().replace(',', ''))
                total += cost
            except (ValueError, AttributeError):
                pass

        # 小計を更新
        self.guest_subtotal.set(f" {int(total):,}")

    def update_guest_arch_costs(self, *args):
        """客室タブの建築項目の金額を自動計算（客室数を反映）"""
        # 和室(10-15J)
        try:
            area = float(self.guest_arch_japanese_area.get()) if self.guest_arch_japanese_area.get() else 0
            room_count = int(self.japanese_room_count.get()) if self.japanese_room_count.get() else 0

            # 部屋数ラベルを更新
            self.guest_japanese_room_count_label.config(text=f"({room_count}室)")

            if area > 0 and room_count > 0:
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("和室(10-15J)")

                if height and floor_cost and ceiling_cost and wall_cost:
                    # 1部屋あたりの金額
                    cost_per_room = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    # 客室数を掛ける
                    total_cost = cost_per_room * room_count
                    self.guest_arch_japanese_cost.set(f"{int(total_cost):,}")

                    process = (f"【和室(10-15J) - {grade_name}】\n"
                               f"1部屋あたりの金額:\n"
                               f"  床: {area}㎡ × {floor_cost:,}円/㎡ = {area * floor_cost:,.0f}円\n"
                               f"  天井: {area}㎡ × {ceiling_cost:,}円/㎡ = {area * ceiling_cost:,.0f}円\n"
                               f"  壁: {area:.2f}^0.5 × 4 × {height}m × {wall_cost:,}円/㎡ = {(area ** 0.5) * 4 * height * wall_cost:,.0f}円\n"
                               f"  1部屋合計 = {int(cost_per_room):,}円\n"
                               f"\n【部屋数を反映】\n"
                               f"  {int(cost_per_room):,}円 × {room_count}室 = {int(total_cost):,}円")
                    self.guest_arch_japanese_process.set(process)
                else:
                    self.guest_arch_japanese_cost.set("0")
                    self.guest_arch_japanese_process.set("【和室(10-15J)】単価情報が取得できません")
            else:
                self.guest_arch_japanese_cost.set("0")
                if area == 0 and room_count == 0:
                    self.guest_arch_japanese_process.set("")
                elif area == 0:
                    self.guest_arch_japanese_process.set("【和室(10-15J)】面積を入力してください")
                else:
                    self.guest_arch_japanese_process.set("【和室(10-15J)】コスト計算タブで客室数を入力してください")
        except (ValueError, AttributeError):
            self.guest_arch_japanese_cost.set("0")
            self.guest_arch_japanese_process.set("【和室(10-15J)】入力値が不正です")
            self.guest_japanese_room_count_label.config(text="(0室)")

        # 和室・洋室
        try:
            area = float(
                self.guest_arch_japanese_western_area.get()) if self.guest_arch_japanese_western_area.get() else 0
            room_count = int(self.japanese_western_room_count.get()) if self.japanese_western_room_count.get() else 0

            # 部屋数ラベルを更新
            self.guest_japanese_western_room_count_label.config(text=f"({room_count}室)")

            if area > 0 and room_count > 0:
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("和室・洋室")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost_per_room = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    total_cost = cost_per_room * room_count
                    self.guest_arch_japanese_western_cost.set(f"{int(total_cost):,}")

                    process = (f"【和室・洋室 - {grade_name}】\n"
                               f"1部屋あたりの金額:\n"
                               f"  床: {area}㎡ × {floor_cost:,}円/㎡ = {area * floor_cost:,.0f}円\n"
                               f"  天井: {area}㎡ × {ceiling_cost:,}円/㎡ = {area * ceiling_cost:,.0f}円\n"
                               f"  壁: {area:.2f}^0.5 × 4 × {height}m × {wall_cost:,}円/㎡ = {(area ** 0.5) * 4 * height * wall_cost:,.0f}円\n"
                               f"  1部屋合計 = {int(cost_per_room):,}円\n"
                               f"\n【部屋数を反映】\n"
                               f"  {int(cost_per_room):,}円 × {room_count}室 = {int(total_cost):,}円")
                    self.guest_arch_japanese_western_process.set(process)
                else:
                    self.guest_arch_japanese_western_cost.set("0")
                    self.guest_arch_japanese_western_process.set("【和室・洋室】単価情報が取得できません")
            else:
                self.guest_arch_japanese_western_cost.set("0")
                if area == 0 and room_count == 0:
                    self.guest_arch_japanese_western_process.set("")
                elif area == 0:
                    self.guest_arch_japanese_western_process.set("【和室・洋室】面積を入力してください")
                else:
                    self.guest_arch_japanese_western_process.set("【和室・洋室】コスト計算タブで客室数を入力してください")
        except (ValueError, AttributeError):
            self.guest_arch_japanese_western_cost.set("0")
            self.guest_arch_japanese_western_process.set("【和室・洋室】入力値が不正です")
            self.guest_japanese_western_room_count_label.config(text="(0室)")

        # 和ベッド
        try:
            area = float(self.guest_arch_japanese_bed_area.get()) if self.guest_arch_japanese_bed_area.get() else 0
            room_count = int(self.japanese_bed_room_count.get()) if self.japanese_bed_room_count.get() else 0

            # 部屋数ラベルを更新
            self.guest_japanese_bed_room_count_label.config(text=f"({room_count}室)")

            if area > 0 and room_count > 0:
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("和ベッド")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost_per_room = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    total_cost = cost_per_room * room_count
                    self.guest_arch_japanese_bed_cost.set(f"{int(total_cost):,}")

                    process = (f"【和ベッド - {grade_name}】\n"
                               f"1部屋あたりの金額:\n"
                               f"  床: {area}㎡ × {floor_cost:,}円/㎡ = {area * floor_cost:,.0f}円\n"
                               f"  天井: {area}㎡ × {ceiling_cost:,}円/㎡ = {area * ceiling_cost:,.0f}円\n"
                               f"  壁: {area:.2f}^0.5 × 4 × {height}m × {wall_cost:,}円/㎡ = {(area ** 0.5) * 4 * height * wall_cost:,.0f}円\n"
                               f"  1部屋合計 = {int(cost_per_room):,}円\n"
                               f"\n【部屋数を反映】\n"
                               f"  {int(cost_per_room):,}円 × {room_count}室 = {int(total_cost):,}円")
                    self.guest_arch_japanese_bed_process.set(process)
                else:
                    self.guest_arch_japanese_bed_cost.set("0")
                    self.guest_arch_japanese_bed_process.set("【和ベッド】単価情報が取得できません")
            else:
                self.guest_arch_japanese_bed_cost.set("0")
                if area == 0 and room_count == 0:
                    self.guest_arch_japanese_bed_process.set("")
                elif area == 0:
                    self.guest_arch_japanese_bed_process.set("【和ベッド】面積を入力してください")
                else:
                    self.guest_arch_japanese_bed_process.set("【和ベッド】コスト計算タブで客室数を入力してください")
        except (ValueError, AttributeError):
            self.guest_arch_japanese_bed_cost.set("0")
            self.guest_arch_japanese_bed_process.set("【和ベッド】入力値が不正です")
            self.guest_japanese_bed_room_count_label.config(text="(0室)")

        # 洋室
        try:
            area = float(self.guest_arch_western_area.get()) if self.guest_arch_western_area.get() else 0
            room_count = int(self.western_room_count.get()) if self.western_room_count.get() else 0

            # 部屋数ラベルを更新
            self.guest_western_room_count_label.config(text=f"({room_count}室)")

            if area > 0 and room_count > 0:
                height, floor_cost, ceiling_cost, wall_cost, grade_name = self.get_lounge_unit_costs("洋室")

                if height and floor_cost and ceiling_cost and wall_cost:
                    cost_per_room = area * floor_cost + area * ceiling_cost + (area ** 0.5) * 4 * height * wall_cost
                    total_cost = cost_per_room * room_count
                    self.guest_arch_western_cost.set(f"{int(total_cost):,}")

                    process = (f"【洋室 - {grade_name}】\n"
                               f"1部屋あたりの金額:\n"
                               f"  床: {area}㎡ × {floor_cost:,}円/㎡ = {area * floor_cost:,.0f}円\n"
                               f"  天井: {area}㎡ × {ceiling_cost:,}円/㎡ = {area * ceiling_cost:,.0f}円\n"
                               f"  壁: {area:.2f}^0.5 × 4 × {height}m × {wall_cost:,}円/㎡ = {(area ** 0.5) * 4 * height * wall_cost:,.0f}円\n"
                               f"  1部屋合計 = {int(cost_per_room):,}円\n"
                               f"\n【部屋数を反映】\n"
                               f"  {int(cost_per_room):,}円 × {room_count}室 = {int(total_cost):,}円")
                    self.guest_arch_western_process.set(process)
                else:
                    self.guest_arch_western_cost.set("0")
                    self.guest_arch_western_process.set("【洋室】単価情報が取得できません")
            else:
                self.guest_arch_western_cost.set("0")
                if area == 0 and room_count == 0:
                    self.guest_arch_western_process.set("")
                elif area == 0:
                    self.guest_arch_western_process.set("【洋室】面積を入力してください")
                else:
                    self.guest_arch_western_process.set("【洋室】コスト計算タブで客室数を入力してください")
        except (ValueError, AttributeError):
            self.guest_arch_western_cost.set("0")
            self.guest_arch_western_process.set("【洋室】入力値が不正です")
            self.guest_western_room_count_label.config(text="(0室)")

        self.update_guest_recommended_counts()
        self.update_guest_subtotal()

    def get_total_guests(self):
        """客室数から合計宿泊人数を計算する"""
        total_guests = 0
        try:
            cap_japanese = float(self.capacity_japanese_room.get())
            cap_japanese_western = float(self.capacity_japanese_western_room.get())
            cap_japanese_bed = float(self.capacity_japanese_bed_room.get())
            cap_western = float(self.capacity_western_room.get())

            total_guests += int(self.japanese_room_count.get()) * cap_japanese
            total_guests += int(self.japanese_western_room_count.get()) * cap_japanese_western
            total_guests += int(self.japanese_bed_room_count.get()) * cap_japanese_bed
            total_guests += int(self.western_room_count.get()) * cap_western
        except ValueError:
            return 0
        return total_guests

    def update_total_guests_realtime(self, *args):
        """リアルタイムで合計宿泊人数を更新する"""
        total = self.get_total_guests()
        self.total_guests_label.config(text=f"合計宿泊人数: {int(total)} 人")
        self.update_onsen_calculations()

    def update_guest_room_capacities(self, *args):
        """客室設定タブの定員が変更されたときに、コスト計算タブの表示も更新する"""
        try:
            cap_japanese = self.capacity_japanese_room.get()
            cap_japanese_western = self.capacity_japanese_western_room.get()
            cap_japanese_bed = self.capacity_japanese_bed_room.get()
            cap_western = self.capacity_western_room.get()

            self.japanese_capacity_label.config(text=f"{cap_japanese}人/室")
            self.japanese_western_capacity_label.config(text=f"{cap_japanese_western}人/室")
            self.japanese_bed_capacity_label.config(text=f"{cap_japanese_bed}人/室")
            self.western_capacity_label.config(text=f"{cap_western}人/室")
        except:
            pass

        self.update_total_guests_realtime()
    # 客室のタブに関する項目終了

    # 建築初期設定のタブに関する項目開始
    def create_structure_unit_cost_table(self, parent_frame, start_row):
        """建築構造別単価設定の表を作成する"""
        current_row = start_row

        # 建物構造設定
        ttk.Label(parent_frame, text="建物構造別設定", font=sub_title_font).grid(
            row=current_row, column=0, columnspan=6, pady=(15, 5), sticky=tk.W
        )
        current_row += 1

        # ヘッダー
        header_labels = ["構造", "Gensen", "Premium1", "Premium2", "TAOYA1", "TAOYA2"]

        for col, text in enumerate(header_labels):
            ttk.Label(parent_frame, text=text, font=sheet_font).grid(
                row=current_row, column=col, sticky=tk.W, padx=5
            )

        current_row += 1

        # データ行 (行: 木造, RC造, 鉄骨造, SRC造)
        for structure, cost_vars in self.structure_unit_costs:
            # 構造名 (行ヘッダー)
            ttk.Label(parent_frame, text=structure, font=sheet_font).grid(
                row=current_row, column=0, sticky=tk.W, padx=5
            )
            # 単価入力欄 (データセル)
            for col in range(5):
                ttk.Entry(parent_frame, textvariable=cost_vars[col], width=10, justify=tk.RIGHT, font=sheet_font).grid(
                    row=current_row, column=col + 1, sticky=(tk.W, tk.E), padx=2, pady=1
                )
            current_row += 1

        # 内装単価設定テーブルの開始行を返す
        return current_row

    def save_arch_data(self):
        """建築内装単価データをJSONファイルに保存"""
        data_to_save = []
        for name, height_var, cost_vars in self.arch_unit_costs:
            try:
                height = float(height_var.get())
                costs = [float(var.get()) for var in cost_vars]
                data_to_save.append([name, height] + costs)
            except ValueError:
                continue

        try:
            with open(self.arch_data_file, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("保存完了", "建築内装単価データが保存されました。")

            #単価設定を保存した時に再計算
            self.update_all_calculations()


        except Exception as e:
            messagebox.showerror("保存エラー", f"データの保存に失敗しました:\n{str(e)}")

    def load_arch_data(self):
        """建築内装単価データをJSONファイルから読み込む"""
        if os.path.exists(self.arch_data_file):
            try:
                with open(self.arch_data_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return data
            except Exception as e:
                messagebox.showwarning("読み込みエラー",
                                       f"データの読み込みに失敗しました:\n{str(e)}\n初期データを使用します。")
                return None
        return None

    def reset_arch_data(self):
        """建築内装単価データを初期状態にリセット"""
        result = messagebox.askyesno("確認",
                                     "建築内装単価データを初期状態にリセットしますか?\n追加したデータは削除されます。")
        if result:
            self.arch_unit_costs.clear()

            for name, height, *costs in self.initial_unit_cost_data:
                height_var = tk.StringVar(value=str(height))
                cost_vars = [tk.StringVar(value=str(c)) for c in costs]
                self.arch_unit_costs.append((name, height_var, cost_vars))

            if os.path.exists(self.arch_data_file):
                try:
                    os.remove(self.arch_data_file)
                except:
                    pass

            self.refresh_arch_settings_tab()
            messagebox.showinfo("リセット完了", "建築内装単価データが初期状態にリセットされました。")

    def add_arch_row(self):
        """建築内装単価の表に新しい行を追加"""
        name = f"新規用途{len(self.arch_unit_costs) + 1}"
        height_var = tk.StringVar(value="2.5")
        cost_vars = [tk.StringVar(value="0") for _ in range(15)]

        self.arch_unit_costs.append((name, height_var, cost_vars))
        self.refresh_arch_settings_tab()

    def delete_arch_row(self):
        """建築内装単価の表から最後の行を削除"""
        if len(self.arch_unit_costs) > 0:
            self.arch_unit_costs.pop()
            self.refresh_arch_settings_tab()

    def refresh_arch_settings_tab(self):
        """建築初期設定タブの表示を更新"""
        for widget in self.arch_settings_frame.winfo_children():
            widget.destroy()
        self.create_arch_settings_tab()
    # 建築初期設定のタブに関する項目終了

    # 設備設定のタブに関する項目開始
    def save_equipment_data(self):
        """設備機器データをJSONファイルに保存"""
        data_to_save = []
        for data_vars in self.equipment_unit_data:
            no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var, new_equipment_cost_var, area_ratio_var = data_vars

            no = no_var.get() if no_var.get() != "" else None
            name = name_var.get() if name_var.get() != "" else None
            abbrev = abbrev_var.get() if abbrev_var.get() != "" else None
            install_hours = install_hours_var.get() if install_hours_var.get() != "" else None
            labor_cost = labor_cost_var.get() if labor_cost_var.get() != "" else None
            misc_material = misc_material_var.get() if misc_material_var.get() != "" else None
            new_equipment_cost = new_equipment_cost_var.get() if new_equipment_cost_var.get() != "" else None
            area_ratio = area_ratio_var.get() if area_ratio_var.get() != "" else None

            data_to_save.append(
                [no, name, abbrev, install_hours, labor_cost, misc_material, new_equipment_cost, area_ratio])

        try:
            with open(self.equipment_data_file, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, ensure_ascii=False, indent=2)
            # 保存後に各タブの計算を再実行
            self.recalculate_all_tabs_after_equipment_change()
            messagebox.showinfo("保存完了", "設備機器データが保存されました。")
        except Exception as e:
            messagebox.showerror("保存エラー", f"データの保存に失敗しました:\n{str(e)}")

    def load_equipment_data(self):
        """設備機器データをJSONファイルから読み込む"""
        if os.path.exists(self.equipment_data_file):
            try:
                with open(self.equipment_data_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return data
            except Exception as e:
                messagebox.showwarning("読み込みエラー",
                                       f"データの読み込みに失敗しました:\n{str(e)}\n初期データを使用します。")
                return None
        return None

    def reset_equipment_data(self):
        """設備機器データを初期状態にリセット"""
        result = messagebox.askyesno("確認",
                                     "設備機器データを初期状態にリセットしますか?\n追加したデータは削除されます。")
        if result:
            self.equipment_unit_data.clear()

            for data in self.initial_equipment_data:
                no, name, abbrev, install_hours, labor_cost, misc_material, new_equipment_cost, area_ratio = data

                no_var = tk.StringVar(value=str(no) if no is not None else "")
                name_var = tk.StringVar(value=name if name is not None else "")
                abbrev_var = tk.StringVar(value=abbrev if abbrev is not None else "")
                install_hours_var = tk.StringVar(value=str(install_hours) if install_hours is not None else "")
                labor_cost_var = tk.StringVar(value=str(labor_cost) if labor_cost is not None else "")
                misc_material_var = tk.StringVar(value=str(misc_material) if misc_material is not None else "")
                new_equipment_cost_var = tk.StringVar(
                    value=str(new_equipment_cost) if new_equipment_cost is not None else "")
                area_ratio_var = tk.StringVar(value=str(area_ratio) if area_ratio is not None else "")

                self.equipment_unit_data.append((
                    no_var, name_var, abbrev_var, install_hours_var,
                    labor_cost_var, misc_material_var, new_equipment_cost_var, area_ratio_var
                ))

            if os.path.exists(self.equipment_data_file):
                try:
                    os.remove(self.equipment_data_file)
                except:
                    pass

            self.refresh_equipment_tab()
            messagebox.showinfo("リセット完了", "設備機器データが初期状態にリセットされました。")

    def add_equipment_row(self):
        """設備機器単価の表に新しい行を追加"""
        new_no = len(self.equipment_unit_data) + 1

        no_var = tk.StringVar(value=str(new_no))
        name_var = tk.StringVar(value="")
        abbrev_var = tk.StringVar(value="")
        install_hours_var = tk.StringVar(value="")
        labor_cost_var = tk.StringVar(value="27000")
        misc_material_var = tk.StringVar(value="")
        new_equipment_cost_var = tk.StringVar(value="")
        area_ratio_var = tk.StringVar(value="")

        self.equipment_unit_data.append((
            no_var, name_var, abbrev_var, install_hours_var,
            labor_cost_var, misc_material_var, new_equipment_cost_var, area_ratio_var
        ))

        self.refresh_equipment_tab()

    def delete_equipment_row(self):
        """設備機器単価の表から最後の行を削除"""
        if len(self.equipment_unit_data) > 0:
            self.equipment_unit_data.pop()
            for i, data_vars in enumerate(self.equipment_unit_data):
                data_vars[0].set(str(i + 1))
            self.refresh_equipment_tab()

    def refresh_equipment_tab(self):
        """設備設定タブの表示を更新"""
        for widget in self.equipment_frame.winfo_children():
            widget.destroy()
        self.create_equipment_tab()

    def create_equipment_tab(self):
        """設備設定タブの作成"""
        canvas = tk.Canvas(self.equipment_frame)
        scrollbar = ttk.Scrollbar(self.equipment_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        row = 0

        ttk.Label(main_frame, text="設備設定", font=title_font).grid(row=row, column=0, columnspan=9,
                                                                                pady=(0, 20), sticky=tk.W)
        row += 1

        title_button_frame = ttk.Frame(main_frame)
        title_button_frame.grid(row=row, column=0, columnspan=9, pady=(10, 10), sticky=tk.W)

        ttk.Label(title_button_frame, text="設備機器単価", font=sub_title_font).pack(side=tk.LEFT)

        ttk.Button(title_button_frame, text="追加", command=self.add_equipment_row).pack(side=tk.LEFT, padx=(10, 5))
        ttk.Button(title_button_frame, text="削除", command=self.delete_equipment_row).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(title_button_frame, text="保存", command=self.save_equipment_data).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(title_button_frame, text="リセット", command=self.reset_equipment_data).pack(side=tk.LEFT)

        row += 1

        headers = ["No", "部材", "略称", "取付工数", "労務単価", "雑材料", "新設機器単価", "面積当り個数"]

        for col, header_text in enumerate(headers):
            ttk.Label(main_frame, text=header_text, font=sheet_font, anchor="center").grid(row=row,
                                                                                                      column=col,
                                                                                                      sticky="ew",
                                                                                                      padx=1, pady=1)

        row += 1

        for data_vars in self.equipment_unit_data:
            no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var, new_equipment_cost_var, area_ratio_var = data_vars

            ttk.Entry(main_frame, textvariable=no_var, width=no_width, justify=tk.CENTER, state='readonly').grid(row=row,
                                                                                                          column=0,
                                                                                                          sticky="ew",
                                                                                                          padx=1,
                                                                                                          pady=1)
            ttk.Entry(main_frame, textvariable=name_var, width=component_name_width, justify=tk.LEFT).grid(row=row, column=1, sticky="ew",
                                                                                         padx=1, pady=1)
            ttk.Entry(main_frame, textvariable=abbrev_var, width=short_name_width, justify=tk.LEFT).grid(row=row, column=2,
                                                                                           sticky="ew", padx=1, pady=1)
            ttk.Entry(main_frame, textvariable=install_hours_var, width=man_day_rate_width, justify=tk.RIGHT).grid(row=row, column=3,
                                                                                                   sticky="ew", padx=1,
                                                                                                   pady=1)
            ttk.Entry(main_frame, textvariable=labor_cost_var, width=labor_cost_width, justify=tk.RIGHT).grid(row=row, column=4,
                                                                                                sticky="ew", padx=1,
                                                                                                pady=1)
            ttk.Entry(main_frame, textvariable=misc_material_var, width=sub_materials_width, justify=tk.RIGHT).grid(row=row, column=5,
                                                                                                   sticky="ew", padx=1,
                                                                                                   pady=1)
            ttk.Entry(main_frame, textvariable=new_equipment_cost_var, width=unit_cost_width, justify=tk.RIGHT).grid(row=row,
                                                                                                        column=6,
                                                                                                        sticky="ew",
                                                                                                        padx=1, pady=1)
            ttk.Entry(main_frame, textvariable=area_ratio_var, width=per_unit_area_width, justify=tk.RIGHT).grid(row=row, column=7,
                                                                                                sticky="ew", padx=1,
                                                                                                pady=1)

            row += 1

        for col in range(8):
            main_frame.grid_columnconfigure(col, weight=1)

        row += 2

        # ========== 熱源機器設定 ==========
        heat_title_button_frame = ttk.Frame(main_frame)#フレームの配置
        heat_title_button_frame.grid(row=row, column=0, columnspan=9, pady=(20, 10), sticky=tk.W)

        ttk.Label(heat_title_button_frame, text="熱源機器設定", font=sub_title_font).pack(side=tk.LEFT)
        ttk.Button(heat_title_button_frame, text="追加", command=self.add_heat_source_row).pack(side=tk.LEFT,
                                                                                                padx=(10, 5))
        ttk.Button(heat_title_button_frame, text="削除", command=self.delete_heat_source_row).pack(side=tk.LEFT,
                                                                                                   padx=(0, 5))
        ttk.Button(heat_title_button_frame, text="保存", command=self.save_heat_source_data).pack(side=tk.LEFT,
                                                                                                  padx=(0, 5))
        ttk.Button(heat_title_button_frame, text="リセット", command=self.reset_heat_source_data).pack(side=tk.LEFT)

        row += 1

        headers_heat = ["No", "部材", "略称", "取付工数", "労務単価", "雑材料", "新設機器単価", "面積当り個数",
                        "燃料種別", "出力"]  # ← "出力"追加
        for col, header_text in enumerate(headers_heat):
            ttk.Label(main_frame, text=header_text, font=sheet_font, anchor="center").grid(row=row,
                                                                                           column=col,
                                                                                           sticky="ew",
                                                                                           padx=1, pady=1)

        row += 1

        for data_vars in self.heat_source_unit_data:
            no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var, new_equipment_cost_var, area_ratio_var, fuel_type_var, output_var = data_vars  # ← output_var追加

            ttk.Entry(main_frame, textvariable=no_var, width=no_width, justify=tk.CENTER, state='readonly').grid(
                row=row,
                column=0,
                sticky="ew",
                padx=1,
                pady=1)
            ttk.Entry(main_frame, textvariable=name_var, width=component_name_width, justify=tk.LEFT).grid(row=row,
                                                                                                           column=1,
                                                                                                           sticky="ew",
                                                                                                           padx=1,
                                                                                                           pady=1)
            ttk.Entry(main_frame, textvariable=abbrev_var, width=short_name_width, justify=tk.LEFT).grid(row=row,
                                                                                                         column=2,
                                                                                                         sticky="ew",
                                                                                                         padx=1, pady=1)
            ttk.Entry(main_frame, textvariable=install_hours_var, width=man_day_rate_width, justify=tk.RIGHT).grid(
                row=row, column=3,
                sticky="ew", padx=1,
                pady=1)
            ttk.Entry(main_frame, textvariable=labor_cost_var, width=labor_cost_width, justify=tk.RIGHT).grid(row=row,
                                                                                                              column=4,
                                                                                                              sticky="ew",
                                                                                                              padx=1,
                                                                                                              pady=1)
            ttk.Entry(main_frame, textvariable=misc_material_var, width=sub_materials_width, justify=tk.RIGHT).grid(
                row=row, column=5,
                sticky="ew", padx=1,
                pady=1)
            ttk.Entry(main_frame, textvariable=new_equipment_cost_var, width=unit_cost_width, justify=tk.RIGHT).grid(
                row=row,
                column=6,
                sticky="ew",
                padx=1, pady=1)
            ttk.Entry(main_frame, textvariable=area_ratio_var, width=per_unit_area_width, justify=tk.RIGHT).grid(
                row=row, column=7,
                sticky="ew", padx=1,
                pady=1)
            ttk.Entry(main_frame, textvariable=fuel_type_var, width=others_width, justify=tk.LEFT).grid(row=row,
                                                                                                        column=8,
                                                                                                        sticky="ew",
                                                                                                        padx=1,
                                                                                                        pady=1)
            ttk.Entry(main_frame, textvariable=output_var, width=others_width, justify=tk.LEFT).grid(row=row, column=9,
                                                                                                     # ← 新規追加
                                                                                                     sticky="ew",
                                                                                                     padx=1,
                                                                                                     pady=1)

            row += 1
        row += 1

        # ========== 厨房機器設定 ==========
        kitchen_title_button_frame = ttk.Frame(main_frame)
        kitchen_title_button_frame.grid(row=row, column=0, columnspan=10, pady=(20, 10), sticky=tk.W)

        ttk.Label(kitchen_title_button_frame, text="厨房機器設定", font=sub_title_font).pack(side=tk.LEFT)
        ttk.Button(kitchen_title_button_frame, text="追加", command=self.add_kitchen_equipment_row).pack(side=tk.LEFT,
                                                                                                         padx=(10, 5))
        ttk.Button(kitchen_title_button_frame, text="削除", command=self.delete_kitchen_equipment_row).pack(
            side=tk.LEFT, padx=(0, 5))
        ttk.Button(kitchen_title_button_frame, text="保存", command=self.save_kitchen_equipment_data).pack(side=tk.LEFT,
                                                                                                           padx=(0, 5))
        ttk.Button(kitchen_title_button_frame, text="リセット", command=self.reset_kitchen_equipment_data).pack(
            side=tk.LEFT)

        row += 1

        # ヘッダー行
        headers_kitchen = ["No", "機器名称", "略称", "取付工数", "労務単価", "経費", "定価", "掛け率", "機器単価",
                           "単位入力", "区分"]
        for col, header_text in enumerate(headers_kitchen):
            ttk.Label(main_frame, text=header_text, font=sheet_font, anchor="center").grid(row=row,
                                                                                                      column=col,
                                                                                                      sticky="ew",
                                                                                                      padx=1, pady=1)

        row += 1

        # データ行の表示（計算結果表示用の変数も保持）
        self.kitchen_cost_display_vars = []  # 計算結果表示用の変数リストを追加

        for idx, data_vars in enumerate(self.kitchen_equipment_unit_data):
            no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, expense_var, list_price_var, rate_var, unit_input_var, category_var = data_vars

            ttk.Entry(main_frame, textvariable=no_var, width=no_width, justify=tk.CENTER, state='readonly').grid(row=row,
                                                                                                          column=0,
                                                                                                          sticky="ew",
                                                                                                          padx=1,
                                                                                                          pady=1)
            ttk.Entry(main_frame, textvariable=name_var, width=component_name_width, justify=tk.LEFT).grid(row=row, column=1, sticky="ew",
                                                                                         padx=1, pady=1)
            ttk.Entry(main_frame, textvariable=abbrev_var, width=short_name_width, justify=tk.LEFT).grid(row=row, column=2,
                                                                                           sticky="ew", padx=1, pady=1)
            ttk.Entry(main_frame, textvariable=install_hours_var, width=man_day_rate_width, justify=tk.RIGHT).grid(row=row, column=3,
                                                                                                   sticky="ew", padx=1,
                                                                                                   pady=1)
            ttk.Entry(main_frame, textvariable=labor_cost_var, width=labor_cost_width, justify=tk.RIGHT).grid(row=row, column=4,
                                                                                                sticky="ew", padx=1,
                                                                                                pady=1)
            ttk.Entry(main_frame, textvariable=expense_var, width=sub_materials_width, justify=tk.RIGHT).grid(row=row, column=5,
                                                                                             sticky="ew", padx=1,
                                                                                             pady=1)
            ttk.Entry(main_frame, textvariable=list_price_var, width=unit_cost_width, justify=tk.RIGHT).grid(row=row, column=6,
                                                                                                sticky="ew", padx=1,
                                                                                                pady=1)
            ttk.Entry(main_frame, textvariable=rate_var, width=price_rate_width, justify=tk.RIGHT).grid(row=row, column=7, sticky="ew",
                                                                                         padx=1, pady=1)

            # 新設機器単価の初期計算
            calculated_cost = self.calculate_kitchen_equipment_cost(
                install_hours_var.get(), labor_cost_var.get(), expense_var.get(),
                list_price_var.get(), rate_var.get()
            )
            cost_display_var = tk.StringVar(value=f"{calculated_cost:,.0f}" if calculated_cost else "0")
            self.kitchen_cost_display_vars.append(cost_display_var)  # リストに保存

            ttk.Entry(main_frame, textvariable=cost_display_var, width=unit_cost_width, justify=tk.RIGHT,
                      state='readonly', foreground='blue').grid(row=row, column=8, sticky="ew", padx=1, pady=1)

            ttk.Entry(main_frame, textvariable=unit_input_var, width=others_width, justify=tk.LEFT).grid(row=row, column=9,
                                                                                               sticky="ew", padx=1,
                                                                                               pady=1)
            ttk.Entry(main_frame, textvariable=category_var, width=others_width, justify=tk.LEFT).grid(row=row, column=10,
                                                                                            sticky="ew", padx=1, pady=1)

            # 各入力項目の変更を監視して自動計算
            install_hours_var.trace('w', lambda *args, i=idx: self.update_kitchen_equipment_cost(i))
            labor_cost_var.trace('w', lambda *args, i=idx: self.update_kitchen_equipment_cost(i))
            expense_var.trace('w', lambda *args, i=idx: self.update_kitchen_equipment_cost(i))
            list_price_var.trace('w', lambda *args, i=idx: self.update_kitchen_equipment_cost(i))
            rate_var.trace('w', lambda *args, i=idx: self.update_kitchen_equipment_cost(i))

            row += 1
        row += 1


        ttk.Label(main_frame, text="受変電設備設定", font=sub_title_font).grid(row=row, column=0, columnspan=9,
                                                                                      pady=(20, 10), sticky=tk.W)
        row += 1
        ttk.Label(main_frame, text="現在開発中です。", font=item_font).grid(row=row, column=0, columnspan=9, pady=10,
                                                                               sticky=tk.W)
        row += 2

        ttk.Label(main_frame, text="中央監視装置設定", font=sub_title_font).grid(row=row, column=0, columnspan=9,
                                                                                        pady=(20, 10), sticky=tk.W)
        row += 1
        ttk.Label(main_frame, text="現在開発中です。", font=item_font).grid(row=row, column=0, columnspan=9, pady=10,
                                                                               sticky=tk.W)

    def calculate_kitchen_equipment_cost(self, install_hours, labor_cost, expense, list_price, rate):
        """
        厨房機器の新設機器単価を計算
        新設機器単価 = 取付工数×労務単価 + 定価×掛け率 + 経費
        """
        try:
            hours = float(install_hours) if install_hours else 0
            labor = float(labor_cost) if labor_cost else 0
            exp = float(expense) if expense else 0
            price = float(list_price) if list_price else 0
            r = float(rate) if rate else 0

            calculated_cost = hours * labor + price * r + exp
            return calculated_cost
        except (ValueError, TypeError):
            return 0

    def update_kitchen_equipment_cost(self, index):
        """
        厨房機器の新設機器単価を動的に更新

        Args:
            index (int): 更新する行のインデックス
        """
        try:
            if index < len(self.kitchen_equipment_unit_data):
                data_vars = self.kitchen_equipment_unit_data[index]
                no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, expense_var, list_price_var, rate_var, unit_input_var, category_var = data_vars

                # 計算実行
                calculated_cost = self.calculate_kitchen_equipment_cost(
                    install_hours_var.get(),
                    labor_cost_var.get(),
                    expense_var.get(),
                    list_price_var.get(),
                    rate_var.get()
                )

                # 計算結果を表示用変数に設定
                if index < len(self.kitchen_cost_display_vars):
                    self.kitchen_cost_display_vars[index].set(f"{calculated_cost:,.0f}" if calculated_cost else "0")
        except Exception as e:
            # エラーが発生しても処理を継続
            pass

    def load_heat_source_data(self):
        """熱源機器データをJSONファイルから読み込む"""
        if os.path.exists(self.heat_source_data_file):
            try:
                with open(self.heat_source_data_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return data
            except Exception as e:
                messagebox.showwarning("読み込みエラー",
                                       f"データの読み込みに失敗しました:\n{str(e)}\n初期データを使用します。")
                return None
        return None

    def save_heat_source_data(self):
        """熱源機器データをJSONファイルに保存"""
        data_to_save = []
        for data_vars in self.heat_source_unit_data:
            no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var, new_equipment_cost_var, area_ratio_var, fuel_type_var, output_var = data_vars  # ← output_var追加
            no = no_var.get() if no_var.get() != "" else None
            name = name_var.get() if name_var.get() != "" else None
            abbrev = abbrev_var.get() if abbrev_var.get() != "" else None
            install_hours = install_hours_var.get() if install_hours_var.get() != "" else None
            labor_cost = labor_cost_var.get() if labor_cost_var.get() != "" else None
            misc_material = misc_material_var.get() if misc_material_var.get() != "" else None
            new_equipment_cost = new_equipment_cost_var.get() if new_equipment_cost_var.get() != "" else None
            area_ratio = area_ratio_var.get() if area_ratio_var.get() != "" else None
            fuel_type = fuel_type_var.get() if fuel_type_var.get() != "" else None
            output = output_var.get() if output_var.get() != "" else None  # ← 追加
            data_to_save.append(
                [no, name, abbrev, install_hours, labor_cost, misc_material, new_equipment_cost, area_ratio, fuel_type, output])  # ← output追加

        try:
            with open(self.heat_source_data_file, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, ensure_ascii=False, indent=2)

            # 保存後に熱源機器の計算を再実行
            if hasattr(self, 'update_heat_source_equipment'):
                self.update_heat_source_equipment()
            messagebox.showinfo("保存完了", "熱源機器データが保存されました。")
        except Exception as e:
            messagebox.showerror("保存エラー", f"データの保存に失敗しました:\n{str(e)}")

    def reset_heat_source_data(self):
        """熱源機器データを初期状態にリセット"""
        result = messagebox.askyesno("確認",
                                     "熱源機器データを初期状態にリセットしますか?\n追加したデータは削除されます。")
        if result:
            self.heat_source_unit_data.clear()
            for data in self.initial_heat_source_data:
                no, name, abbrev, install_hours, labor_cost, misc_material, new_equipment_cost, area_ratio, fuel_type, output = data  # ← output追加
                no_var = tk.StringVar(value=str(no) if no is not None else "")
                name_var = tk.StringVar(value=name if name is not None else "")
                abbrev_var = tk.StringVar(value=abbrev if abbrev is not None else "")
                install_hours_var = tk.StringVar(value=str(install_hours) if install_hours is not None else "")
                labor_cost_var = tk.StringVar(value=str(labor_cost) if labor_cost is not None else "")
                misc_material_var = tk.StringVar(value=str(misc_material) if misc_material is not None else "")
                new_equipment_cost_var = tk.StringVar(
                    value=str(new_equipment_cost) if new_equipment_cost is not None else "")
                area_ratio_var = tk.StringVar(value=str(area_ratio) if area_ratio is not None else "")
                fuel_type_var = tk.StringVar(value=str(fuel_type) if fuel_type is not None else "")
                output_var = tk.StringVar(value=str(output) if output is not None else "")  # ← 追加
                self.heat_source_unit_data.append(
                    (no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var,
                     new_equipment_cost_var, area_ratio_var, fuel_type_var, output_var))  # ← output_var追加

            if os.path.exists(self.heat_source_data_file):
                try:
                    os.remove(self.heat_source_data_file)
                except:
                    pass

            self.refresh_equipment_tab()
            messagebox.showinfo("リセット完了", "熱源機器データが初期状態にリセットされました。")

    def add_heat_source_row(self):
        """熱源機器設定の表に新しい行を追加"""
        new_no = len(self.heat_source_unit_data) + 1
        no_var = tk.StringVar(value=str(new_no))
        name_var = tk.StringVar(value="")
        abbrev_var = tk.StringVar(value="")
        install_hours_var = tk.StringVar(value="0.1")
        labor_cost_var = tk.StringVar(value="27000")
        misc_material_var = tk.StringVar(value="2000")
        new_equipment_cost_var = tk.StringVar(value="30000")
        area_ratio_var = tk.StringVar(value="0.02")
        fuel_type_var = tk.StringVar(value="")
        output_var = tk.StringVar(value="")  # ← 追加

        self.heat_source_unit_data.append((
            no_var, name_var, abbrev_var, install_hours_var,
            labor_cost_var, misc_material_var, new_equipment_cost_var, area_ratio_var, fuel_type_var, output_var  # ← output_var追加
        ))
        self.refresh_equipment_tab()

    def delete_heat_source_row(self):
        """熱源機器設定の表から最後の行を削除"""
        if len(self.heat_source_unit_data) > 0:
            self.heat_source_unit_data.pop()
            for i, data_vars in enumerate(self.heat_source_unit_data):
                data_vars[0].set(str(i + 1))
            self.refresh_equipment_tab()

    def load_kitchen_equipment_data(self):
        """厨房機器データをJSONファイルから読み込む"""
        if os.path.exists(self.kitchen_equipment_data_file):
            try:
                with open(self.kitchen_equipment_data_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return data
            except Exception as e:
                messagebox.showwarning("読み込みエラー",
                                       f"データの読み込みに失敗しました:\n{str(e)}\n初期データを使用します。")
                return None
        return None

    def save_kitchen_equipment_data(self):
        """厨房機器データをJSONファイルに保存"""
        data_to_save = []
        for data_vars in self.kitchen_equipment_unit_data:
            no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, expense_var, list_price_var, rate_var, unit_input_var, category_var = data_vars
            no = no_var.get() if no_var.get() != "" else None
            name = name_var.get() if name_var.get() != "" else None
            abbrev = abbrev_var.get() if abbrev_var.get() != "" else None
            install_hours = install_hours_var.get() if install_hours_var.get() != "" else None
            labor_cost = labor_cost_var.get() if labor_cost_var.get() != "" else None
            expense = expense_var.get() if expense_var.get() != "" else None
            list_price = list_price_var.get() if list_price_var.get() != "" else None
            rate = rate_var.get() if rate_var.get() != "" else None
            unit_input = unit_input_var.get() if unit_input_var.get() != "" else None
            category = category_var.get() if category_var.get() != "" else None
            # 新設機器単価は保存しない（計算で求めるため）
            data_to_save.append(
                [no, name, abbrev, install_hours, labor_cost, expense, list_price, rate, unit_input, category])

        try:
            with open(self.kitchen_equipment_data_file, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, ensure_ascii=False, indent=2)
            # 保存後にレストランタブの厨房機器計算を再実行
            if hasattr(self, 'update_restaurant_kitchen_costs'):
                self.update_restaurant_kitchen_costs()
            messagebox.showinfo("保存完了", "厨房機器データが保存されました。")
        except Exception as e:
            messagebox.showerror("保存エラー", f"データの保存に失敗しました:\n{str(e)}")

    def reset_kitchen_equipment_data(self):
        """厨房機器データを初期状態にリセット"""
        result = messagebox.askyesno("確認",
                                     "厨房機器データを初期状態にリセットしますか?\n追加したデータは削除されます。")
        if result:
            self.kitchen_equipment_unit_data.clear()
            for data in self.initial_kitchen_equipment_data:
                no, name, abbrev, install_hours, labor_cost, expense, list_price, rate, unit_input, category = data
                no_var = tk.StringVar(value=str(no) if no is not None else "")
                name_var = tk.StringVar(value=name if name is not None else "")
                abbrev_var = tk.StringVar(value=abbrev if abbrev is not None else "")
                install_hours_var = tk.StringVar(value=str(install_hours) if install_hours is not None else "")
                labor_cost_var = tk.StringVar(value=str(labor_cost) if labor_cost is not None else "")
                expense_var = tk.StringVar(value=str(expense) if expense is not None else "")
                list_price_var = tk.StringVar(value=str(list_price) if list_price is not None else "")
                rate_var = tk.StringVar(value=str(rate) if rate is not None else "")
                unit_input_var = tk.StringVar(value=str(unit_input) if unit_input is not None else "")
                category_var = tk.StringVar(value=str(category) if category is not None else "")
                self.kitchen_equipment_unit_data.append(
                    (no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, expense_var,
                     list_price_var, rate_var, unit_input_var, category_var))

            if os.path.exists(self.kitchen_equipment_data_file):
                try:
                    os.remove(self.kitchen_equipment_data_file)
                except:
                    pass

            self.refresh_equipment_tab()
            messagebox.showinfo("リセット完了", "厨房機器データが初期状態にリセットされました。")

    def add_kitchen_equipment_row(self):
        """厨房機器設定の表に新しい行を追加"""
        new_no = len(self.kitchen_equipment_unit_data) + 1
        no_var = tk.StringVar(value=str(new_no))
        name_var = tk.StringVar(value="")
        abbrev_var = tk.StringVar(value="")
        install_hours_var = tk.StringVar(value="0.74")
        labor_cost_var = tk.StringVar(value="27000")
        expense_var = tk.StringVar(value="10000")
        list_price_var = tk.StringVar(value="100000")
        rate_var = tk.StringVar(value="0.60")
        unit_input_var = tk.StringVar(value="")
        category_var = tk.StringVar(value="")

        self.kitchen_equipment_unit_data.append((
            no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, expense_var,
            list_price_var, rate_var, unit_input_var, category_var
        ))
        self.refresh_equipment_tab()

    def delete_kitchen_equipment_row(self):
        """厨房機器設定の表から最後の行を削除"""
        if len(self.kitchen_equipment_unit_data) > 0:
            self.kitchen_equipment_unit_data.pop()
            for i, data_vars in enumerate(self.kitchen_equipment_unit_data):
                data_vars[0].set(str(i + 1))
            self.refresh_equipment_tab()

    def recalculate_all_tabs_after_equipment_change(self):
        """設備機器データ変更後に全タブの計算を再実行"""
        try:
            # 温泉設備タブの再計算
            if hasattr(self, 'update_wash_equipment_recommended'):
                self.update_wash_equipment_recommended()
            if hasattr(self, 'update_wash_equipment_costs'):
                self.update_wash_equipment_costs()
            if hasattr(self, 'update_onsen_bath_arch_costs'):
                self.update_onsen_bath_arch_costs()
            if hasattr(self, 'update_onsen_subtotal'):
                self.update_onsen_subtotal()

            # レストランタブの再計算
            if hasattr(self, 'update_restaurant_recommended_counts'):
                self.update_restaurant_recommended_counts()
            if hasattr(self, 'update_restaurant_equipment_costs'):
                self.update_restaurant_equipment_costs()
            if hasattr(self, 'update_restaurant_subtotal'):
                self.update_restaurant_subtotal()

            # ラウンジタブの再計算
            if hasattr(self, 'update_lounge_recommended_counts'):
                self.update_lounge_recommended_counts()
            if hasattr(self, 'update_lounge_equipment_costs'):
                self.update_lounge_equipment_costs()
            if hasattr(self, 'update_lounge_subtotal'):
                self.update_lounge_subtotal()

            # アミューズメントタブの再計算
            if hasattr(self, 'update_amusement_recommended_counts'):
                self.update_amusement_recommended_counts()
            if hasattr(self, 'update_amusement_equipment_costs'):
                self.update_amusement_equipment_costs()
            if hasattr(self, 'update_amusement_subtotal'):
                self.update_amusement_subtotal()

            # 通路タブの再計算
            if hasattr(self, 'update_hallway_recommended_counts'):
                self.update_hallway_recommended_counts()
            if hasattr(self, 'update_hallway_equipment_costs'):
                self.update_hallway_equipment_costs()
            if hasattr(self, 'update_hallway_subtotal'):
                self.update_hallway_subtotal()

            # 客室タブの再計算
            if hasattr(self, 'update_guest_room_equipment_costs'):
                self.update_guest_room_equipment_costs()
            if hasattr(self, 'update_guest_room_subtotal'):
                self.update_guest_room_subtotal()

        except Exception as e:
            print(f"再計算エラー: {e}")


    # 設備設定のタブに関する項目終了

    # 家具のタブに関する項目開始
    def get_furniture1_data_by_abbrev(self, abbrev):
        """略称から家具設定単価1データを取得"""
        for data_vars in self.furniture1_unit_data:
            no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var, new_equipment_cost_var, area_ratio_var = data_vars

            if abbrev_var.get() == abbrev:
                try:
                    install_hours = float(install_hours_var.get()) if install_hours_var.get() else 0
                    labor_cost = float(labor_cost_var.get()) if labor_cost_var.get() else 0
                    misc_material = float(misc_material_var.get()) if misc_material_var.get() else 0
                    new_equipment_cost = float(new_equipment_cost_var.get()) if new_equipment_cost_var.get() else 0
                    area_ratio = float(area_ratio_var.get()) if area_ratio_var.get() else 0

                    return {
                        'install_hours': install_hours,
                        'labor_cost': labor_cost,
                        'misc_material': misc_material,
                        'new_equipment_cost': new_equipment_cost,
                        'area_ratio': area_ratio,
                        'name': name_var.get()
                    }
                except (ValueError, AttributeError):
                    return None
        return None

    def update_restaurant_furniture_costs(self, *args):
        """レストランの家具の金額を計算"""
        furniture_process_text = "【家具の計算】\n"

        # バーカウンター
        try:
            count = float(
                self.restaurant_furniture_bar_counter_count.get()) if self.restaurant_furniture_bar_counter_count.get() else 0
            if count > 0:
                data = self.get_furniture1_data_by_abbrev("バーカウンター")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.restaurant_furniture_bar_counter_cost.set(f"{int(cost):,}")
                    furniture_process_text += f"\n[バーカウンター]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.restaurant_furniture_bar_counter_cost.set("0")
        except (ValueError, TypeError):
            self.restaurant_furniture_bar_counter_cost.set("0")

        # ソフトドリンクカウンター
        try:
            count = float(
                self.restaurant_furniture_soft_drink_counter_count.get()) if self.restaurant_furniture_soft_drink_counter_count.get() else 0
            if count > 0:
                data = self.get_furniture1_data_by_abbrev("ソフトドリンクカウンター")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.restaurant_furniture_soft_drink_counter_cost.set(f"{int(cost):,}")
                    furniture_process_text += f"\n[ソフトドリンクカウンター]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.restaurant_furniture_soft_drink_counter_cost.set("0")
        except (ValueError, TypeError):
            self.restaurant_furniture_soft_drink_counter_cost.set("0")

        # アルコールカウンター
        try:
            count = float(
                self.restaurant_furniture_alcohol_counter_count.get()) if self.restaurant_furniture_alcohol_counter_count.get() else 0
            if count > 0:
                data = self.get_furniture1_data_by_abbrev("アルコールカウンター")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.restaurant_furniture_alcohol_counter_cost.set(f"{int(cost):,}")
                    furniture_process_text += f"\n[アルコールカウンター]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.restaurant_furniture_alcohol_counter_cost.set("0")
        except (ValueError, TypeError):
            self.restaurant_furniture_alcohol_counter_cost.set("0")

        # カトラリ―カウンター
        try:
            count = float(
                self.restaurant_furniture_cutlery_counter_count.get()) if self.restaurant_furniture_cutlery_counter_count.get() else 0
            if count > 0:
                data = self.get_furniture1_data_by_abbrev("カトラリ―カウンター")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.restaurant_furniture_cutlery_counter_cost.set(f"{int(cost):,}")
                    furniture_process_text += f"\n[カトラリ―カウンター]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.restaurant_furniture_cutlery_counter_cost.set("0")
        except (ValueError, TypeError):
            self.restaurant_furniture_cutlery_counter_cost.set("0")

        # アイスカウンター
        try:
            count = float(
                self.restaurant_furniture_ice_counter_count.get()) if self.restaurant_furniture_ice_counter_count.get() else 0
            if count > 0:
                data = self.get_furniture1_data_by_abbrev("アイスカウンター")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.restaurant_furniture_ice_counter_cost.set(f"{int(cost):,}")
                    furniture_process_text += f"\n[アイスカウンター]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.restaurant_furniture_ice_counter_cost.set("0")
        except (ValueError, TypeError):
            self.restaurant_furniture_ice_counter_cost.set("0")

        # ソフトクリームカウンター
        try:
            count = float(
                self.restaurant_furniture_soft_cream_counter_count.get()) if self.restaurant_furniture_soft_cream_counter_count.get() else 0
            if count > 0:
                data = self.get_furniture1_data_by_abbrev("ソフトクリームカウンター")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.restaurant_furniture_soft_cream_counter_cost.set(f"{int(cost):,}")
                    furniture_process_text += f"\n[ソフトクリームカウンター]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.restaurant_furniture_soft_cream_counter_cost.set("0")
        except (ValueError, TypeError):
            self.restaurant_furniture_soft_cream_counter_cost.set("0")

        # 返却台
        try:
            count = float(
                self.restaurant_furniture_return_counter_count.get()) if self.restaurant_furniture_return_counter_count.get() else 0
            if count > 0:
                data = self.get_furniture1_data_by_abbrev("返却台")
                if data:
                    cost = (data['install_hours'] * data['labor_cost'] + data['new_equipment_cost'] + data[
                        'misc_material']) / 0.8 * count
                    self.restaurant_furniture_return_counter_cost.set(f"{int(cost):,}")
                    furniture_process_text += f"\n[返却台]\n  金額 = ({data['install_hours']}×{data['labor_cost']:,} + {data['new_equipment_cost']:,} + {data['misc_material']:,}) ÷ 0.8 × {count}\n       = {int(cost):,}円\n"
            else:
                self.restaurant_furniture_return_counter_cost.set("0")
        except (ValueError, TypeError):
            self.restaurant_furniture_return_counter_cost.set("0")

        self.restaurant_furniture_process.set(furniture_process_text)

        # 小計を更新
        self.update_restaurant_subtotal()

    def create_furniture_tab(self):
        """家具タブの作成"""
        canvas = tk.Canvas(self.furniture_frame)
        scrollbar = ttk.Scrollbar(self.furniture_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        row = 0

        ttk.Label(main_frame, text="家具設定", font=title_font).grid(row=row, column=0, columnspan=9, pady=(0, 20),
                                                                     sticky=tk.W)
        row += 1

        # ========== 家具設定単価1 ==========
        title_button_frame1 = ttk.Frame(main_frame)
        title_button_frame1.grid(row=row, column=0, columnspan=9, pady=(10, 10), sticky=tk.W)

        ttk.Label(title_button_frame1, text="家具設定単価1", font=sub_title_font).pack(side=tk.LEFT)
        ttk.Button(title_button_frame1, text="追加", command=self.add_furniture1_row).pack(side=tk.LEFT, padx=(10, 5))
        ttk.Button(title_button_frame1, text="削除", command=self.delete_furniture1_row).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(title_button_frame1, text="保存", command=self.save_furniture1_data).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(title_button_frame1, text="リセット", command=self.reset_furniture1_data).pack(side=tk.LEFT)

        row += 1

        headers1 = ["No", "部材", "略称", "取付工数", "労務単価", "雑材料", "新設機器単価", "単位面積の個数"]
        for col, header_text in enumerate(headers1):
            ttk.Label(main_frame, text=header_text, font=sheet_font, anchor="center").grid(row=row,
                                                                                                      column=col,
                                                                                                      sticky="ew",
                                                                                                      padx=1, pady=1)

        row += 1

        for data_vars in self.furniture1_unit_data:
            no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var, new_equipment_cost_var, area_ratio_var = data_vars

            ttk.Entry(main_frame, textvariable=no_var, width=no_width, justify=tk.CENTER, state='readonly').grid(row=row,
                                                                                                          column=0,
                                                                                                          sticky="ew",
                                                                                                          padx=1,
                                                                                                          pady=1)
            ttk.Entry(main_frame, textvariable=name_var, width=component_name_width, justify=tk.LEFT).grid(row=row, column=1, sticky="ew",
                                                                                         padx=1, pady=1)
            ttk.Entry(main_frame, textvariable=abbrev_var, width=short_name_width, justify=tk.LEFT).grid(row=row, column=2,
                                                                                           sticky="ew", padx=1, pady=1)
            ttk.Entry(main_frame, textvariable=install_hours_var, width=man_day_rate_width, justify=tk.RIGHT).grid(row=row, column=3,
                                                                                                   sticky="ew", padx=1,
                                                                                                   pady=1)
            ttk.Entry(main_frame, textvariable=labor_cost_var, width=labor_cost_width, justify=tk.RIGHT).grid(row=row, column=4,
                                                                                                sticky="ew", padx=1,
                                                                                                pady=1)
            ttk.Entry(main_frame, textvariable=misc_material_var, width=sub_materials_width, justify=tk.RIGHT).grid(row=row, column=5,
                                                                                                   sticky="ew", padx=1,
                                                                                                   pady=1)
            ttk.Entry(main_frame, textvariable=new_equipment_cost_var, width=unit_cost_width, justify=tk.RIGHT).grid(row=row,
                                                                                                        column=6,
                                                                                                        sticky="ew",
                                                                                                        padx=1, pady=1)
            ttk.Entry(main_frame, textvariable=area_ratio_var, width=per_unit_area_width, justify=tk.RIGHT).grid(row=row, column=7,
                                                                                                sticky="ew", padx=1,
                                                                                                pady=1)

            row += 1

        for col in range(8):
            main_frame.grid_columnconfigure(col, weight=1)

        row += 2

        # ========== 家具設定単価2 ==========
        title_button_frame2 = ttk.Frame(main_frame)
        title_button_frame2.grid(row=row, column=0, columnspan=9, pady=(20, 10), sticky=tk.W)

        ttk.Label(title_button_frame2, text="家具設定単価2", font=sub_title_font).pack(side=tk.LEFT)
        ttk.Button(title_button_frame2, text="追加", command=self.add_furniture2_row).pack(side=tk.LEFT, padx=(10, 5))
        ttk.Button(title_button_frame2, text="削除", command=self.delete_furniture2_row).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(title_button_frame2, text="保存", command=self.save_furniture2_data).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(title_button_frame2, text="リセット", command=self.reset_furniture2_data).pack(side=tk.LEFT)

        row += 1

        headers2 = ["No", "部材", "略称", "取付工数", "労務単価", "雑材料", "新設機器単価", "客室当の個数"]
        for col, header_text in enumerate(headers2):
            ttk.Label(main_frame, text=header_text, font=sheet_font, anchor="center").grid(row=row,
                                                                                                      column=col,
                                                                                                      sticky="ew",
                                                                                                      padx=1, pady=1)

        row += 1

        for data_vars in self.furniture2_unit_data:
            no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var, new_equipment_cost_var, room_ratio_var = data_vars

            ttk.Entry(main_frame, textvariable=no_var, width=no_width, justify=tk.CENTER, state='readonly').grid(row=row,
                                                                                                          column=0,
                                                                                                          sticky="ew",
                                                                                                          padx=1,
                                                                                                          pady=1)
            ttk.Entry(main_frame, textvariable=name_var, width=component_name_width, justify=tk.LEFT).grid(row=row, column=1, sticky="ew",
                                                                                         padx=1, pady=1)
            ttk.Entry(main_frame, textvariable=abbrev_var, width=short_name_width, justify=tk.LEFT).grid(row=row, column=2,
                                                                                           sticky="ew", padx=1, pady=1)
            ttk.Entry(main_frame, textvariable=install_hours_var, width=man_day_rate_width, justify=tk.RIGHT).grid(row=row, column=3,
                                                                                                   sticky="ew", padx=1,
                                                                                                   pady=1)
            ttk.Entry(main_frame, textvariable=labor_cost_var, width=labor_cost_width, justify=tk.RIGHT).grid(row=row, column=4,
                                                                                                sticky="ew", padx=1,
                                                                                                pady=1)
            ttk.Entry(main_frame, textvariable=misc_material_var, width=sub_materials_width, justify=tk.RIGHT).grid(row=row, column=5,
                                                                                                   sticky="ew", padx=1,
                                                                                                   pady=1)
            ttk.Entry(main_frame, textvariable=new_equipment_cost_var, width=unit_cost_width, justify=tk.RIGHT).grid(row=row,
                                                                                                        column=6,
                                                                                                        sticky="ew",
                                                                                                        padx=1, pady=1)
            ttk.Entry(main_frame, textvariable=room_ratio_var, width=per_unit_area_width, justify=tk.RIGHT).grid(row=row, column=7,
                                                                                                sticky="ew", padx=1,
                                                                                                pady=1)

            row += 1

    def load_furniture1_data(self):
        """家具設定単価1データをJSONファイルから読み込む"""
        if os.path.exists(self.furniture1_data_file):
            try:
                with open(self.furniture1_data_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return data
            except Exception as e:
                messagebox.showwarning("読み込みエラー",
                                       f"データの読み込みに失敗しました:\n{str(e)}\n初期データを使用します。")
                return None
        return None

    def save_furniture1_data(self):
        """家具設定単価1データをJSONファイルに保存"""
        data_to_save = []
        for data_vars in self.furniture1_unit_data:
            no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var, new_equipment_cost_var, area_ratio_var = data_vars
            no = no_var.get() if no_var.get() != "" else None
            name = name_var.get() if name_var.get() != "" else None
            abbrev = abbrev_var.get() if abbrev_var.get() != "" else None
            install_hours = install_hours_var.get() if install_hours_var.get() != "" else None
            labor_cost = labor_cost_var.get() if labor_cost_var.get() != "" else None
            misc_material = misc_material_var.get() if misc_material_var.get() != "" else None
            new_equipment_cost = new_equipment_cost_var.get() if new_equipment_cost_var.get() != "" else None
            area_ratio = area_ratio_var.get() if area_ratio_var.get() != "" else None
            data_to_save.append(
                [no, name, abbrev, install_hours, labor_cost, misc_material, new_equipment_cost, area_ratio])

        try:
            with open(self.furniture1_data_file, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("保存完了", "家具設定単価1データが保存されました。")
            # 単価設定を保存した時に再計算
            self.update_all_calculations()
        except Exception as e:
            messagebox.showerror("保存エラー", f"データの保存に失敗しました:\n{str(e)}")

    def reset_furniture1_data(self):
        """家具設定単価1データを初期状態にリセット"""
        result = messagebox.askyesno("確認",
                                     "家具設定単価1データを初期状態にリセットしますか?\n追加したデータは削除されます。")
        if result:
            self.furniture1_unit_data.clear()
            for data in self.initial_furniture1_data:
                no, name, abbrev, install_hours, labor_cost, misc_material, new_equipment_cost, area_ratio = data
                no_var = tk.StringVar(value=str(no) if no is not None else "")
                name_var = tk.StringVar(value=name if name is not None else "")
                abbrev_var = tk.StringVar(value=abbrev if abbrev is not None else "")
                install_hours_var = tk.StringVar(value=str(install_hours) if install_hours is not None else "")
                labor_cost_var = tk.StringVar(value=str(labor_cost) if labor_cost is not None else "")
                misc_material_var = tk.StringVar(value=str(misc_material) if misc_material is not None else "")
                new_equipment_cost_var = tk.StringVar(
                    value=str(new_equipment_cost) if new_equipment_cost is not None else "")
                area_ratio_var = tk.StringVar(value=str(area_ratio) if area_ratio is not None else "")
                self.furniture1_unit_data.append(
                    (no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var,
                     new_equipment_cost_var, area_ratio_var))

            if os.path.exists(self.furniture1_data_file):
                try:
                    os.remove(self.furniture1_data_file)
                except:
                    pass

            self.refresh_furniture_tab()
            messagebox.showinfo("リセット完了", "家具設定単価1データが初期状態にリセットされました。")

    def add_furniture1_row(self):
        """家具設定単価1の表に新しい行を追加"""
        new_no = len(self.furniture1_unit_data) + 1
        no_var = tk.StringVar(value=str(new_no))
        name_var = tk.StringVar(value="")
        abbrev_var = tk.StringVar(value="")
        install_hours_var = tk.StringVar(value="0.3")
        labor_cost_var = tk.StringVar(value="27000")
        misc_material_var = tk.StringVar(value="2000")
        new_equipment_cost_var = tk.StringVar(value="120000")
        area_ratio_var = tk.StringVar(value="0.02")

        self.furniture1_unit_data.append(
            (no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var, new_equipment_cost_var,
             area_ratio_var))
        self.refresh_furniture_tab()

    def delete_furniture1_row(self):
        """家具設定単価1の表から最後の行を削除"""
        if len(self.furniture1_unit_data) > 0:
            self.furniture1_unit_data.pop()
            for i, data_vars in enumerate(self.furniture1_unit_data):
                data_vars[0].set(str(i + 1))
            self.refresh_furniture_tab()

    def load_furniture2_data(self):
        """家具設定単価2データをJSONファイルから読み込む"""
        if os.path.exists(self.furniture2_data_file):
            try:
                with open(self.furniture2_data_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return data
            except Exception as e:
                messagebox.showwarning("読み込みエラー",
                                       f"データの読み込みに失敗しました:\n{str(e)}\n初期データを使用します。")
                return None
        return None

    def save_furniture2_data(self):
        """家具設定単価2データをJSONファイルに保存"""
        data_to_save = []
        for data_vars in self.furniture2_unit_data:
            no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var, new_equipment_cost_var, room_ratio_var = data_vars
            no = no_var.get() if no_var.get() != "" else None
            name = name_var.get() if name_var.get() != "" else None
            abbrev = abbrev_var.get() if abbrev_var.get() != "" else None
            install_hours = install_hours_var.get() if install_hours_var.get() != "" else None
            labor_cost = labor_cost_var.get() if labor_cost_var.get() != "" else None
            misc_material = misc_material_var.get() if misc_material_var.get() != "" else None
            new_equipment_cost = new_equipment_cost_var.get() if new_equipment_cost_var.get() != "" else None
            room_ratio = room_ratio_var.get() if room_ratio_var.get() != "" else None
            data_to_save.append(
                [no, name, abbrev, install_hours, labor_cost, misc_material, new_equipment_cost, room_ratio])

        try:
            with open(self.furniture2_data_file, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("保存完了", "家具設定単価2データが保存されました。")
        except Exception as e:
            messagebox.showerror("保存エラー", f"データの保存に失敗しました:\n{str(e)}")

    def reset_furniture2_data(self):
        """家具設定単価2データを初期状態にリセット"""
        result = messagebox.askyesno("確認",
                                     "家具設定単価2データを初期状態にリセットしますか?\n追加したデータは削除されます。")
        if result:
            self.furniture2_unit_data.clear()
            for data in self.initial_furniture2_data:
                no, name, abbrev, install_hours, labor_cost, misc_material, new_equipment_cost, room_ratio = data
                no_var = tk.StringVar(value=str(no) if no is not None else "")
                name_var = tk.StringVar(value=name if name is not None else "")
                abbrev_var = tk.StringVar(value=abbrev if abbrev is not None else "")
                install_hours_var = tk.StringVar(value=str(install_hours) if install_hours is not None else "")
                labor_cost_var = tk.StringVar(value=str(labor_cost) if labor_cost is not None else "")
                misc_material_var = tk.StringVar(value=str(misc_material) if misc_material is not None else "")
                new_equipment_cost_var = tk.StringVar(
                    value=str(new_equipment_cost) if new_equipment_cost is not None else "")
                room_ratio_var = tk.StringVar(value=str(room_ratio) if room_ratio is not None else "")
                self.furniture2_unit_data.append(
                    (no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var,
                     new_equipment_cost_var, room_ratio_var))

            if os.path.exists(self.furniture2_data_file):
                try:
                    os.remove(self.furniture2_data_file)
                except:
                    pass

            self.refresh_furniture_tab()
            messagebox.showinfo("リセット完了", "家具設定単価2データが初期状態にリセットされました。")

    def add_furniture2_row(self):
        """家具設定単価2の表に新しい行を追加"""
        new_no = len(self.furniture2_unit_data) + 1
        no_var = tk.StringVar(value=str(new_no))
        name_var = tk.StringVar(value="")
        abbrev_var = tk.StringVar(value="")
        install_hours_var = tk.StringVar(value="0.3")
        labor_cost_var = tk.StringVar(value="27000")
        misc_material_var = tk.StringVar(value="2000")
        new_equipment_cost_var = tk.StringVar(value="120000")
        room_ratio_var = tk.StringVar(value="0.02")

        self.furniture2_unit_data.append(
            (no_var, name_var, abbrev_var, install_hours_var, labor_cost_var, misc_material_var, new_equipment_cost_var,
             room_ratio_var))
        self.refresh_furniture_tab()

    def delete_furniture2_row(self):
        """家具設定単価2の表から最後の行を削除"""
        if len(self.furniture2_unit_data) > 0:
            self.furniture2_unit_data.pop()
            for i, data_vars in enumerate(self.furniture2_unit_data):
                data_vars[0].set(str(i + 1))
            self.refresh_furniture_tab()

    def refresh_furniture_tab(self):
        """家具タブの表示を更新"""
        for widget in self.furniture_frame.winfo_children():
            widget.destroy()
        self.create_furniture_tab()
    # 家具のタブに関する項目終了

    # 経費設定のタブに関する項目開始
    def create_settings_tab(self):
        """設定タブの作成"""
        settings_main = ttk.Frame(self.settings_frame, padding="20")
        settings_main.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        title_label = ttk.Label(settings_main, text="初期設定", font=title_font)
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 30), sticky=tk.W)

        row = 1

        ttk.Label(settings_main, text="諸経費率設定", font=sub_title_font).grid(row=row, column=0, columnspan=3,
                                                                                pady=(10, 5), sticky=tk.W)
        row += 1

        ttk.Label(settings_main, text="諸経費率 (%):").grid(row=row, column=0, sticky=tk.W, pady=10)
        misc_rate_entry = ttk.Entry(settings_main, textvariable=self.misc_cost_rate, width=10)
        misc_rate_entry.grid(row=row, column=1, sticky=tk.W, pady=10)
        ttk.Label(settings_main, text="(例: 0.15 = 15%)").grid(row=row, column=2, sticky=tk.W, pady=10, padx=(10, 0))
        row += 1

        ttk.Label(settings_main, text="スライダーで調整:").grid(row=row, column=0, sticky=tk.W, pady=5)
        misc_scale = ttk.Scale(settings_main, from_=0.05, to=0.30, orient=tk.HORIZONTAL, variable=self.misc_cost_rate,
                               length=200)
        misc_scale.grid(row=row, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        row += 1

        self.misc_rate_display = ttk.Label(settings_main, text="", font=("Arial", 10))
        self.misc_rate_display.grid(row=row, column=1, sticky=tk.W, pady=5)
        row += 2

        ttk.Label(settings_main, text="仕様別割増率設定", font=sub_title_font).grid(row=row, column=0,
                                                                                    columnspan=3, pady=(20, 5),
                                                                                    sticky=tk.W)
        row += 1

        ttk.Label(settings_main, text="仕様1 (上質) 割増率:").grid(row=row, column=0, sticky=tk.W, pady=8)
        ttk.Entry(settings_main, textvariable=self.spec1_rate, width=10).grid(row=row, column=1, sticky=tk.W, pady=8)
        ttk.Scale(settings_main, from_=0.05, to=0.50, orient=tk.HORIZONTAL, variable=self.spec1_rate, length=150).grid(
            row=row, column=2, sticky=tk.W, pady=8, padx=(10, 0))
        row += 1

        ttk.Label(settings_main, text="仕様2 (高級) 割増率:").grid(row=row, column=0, sticky=tk.W, pady=8)
        ttk.Entry(settings_main, textvariable=self.spec2_rate, width=10).grid(row=row, column=1, sticky=tk.W, pady=8)
        ttk.Scale(settings_main, from_=0.10, to=0.60, orient=tk.HORIZONTAL, variable=self.spec2_rate, length=150).grid(
            row=row, column=2, sticky=tk.W, pady=8, padx=(10, 0))
        row += 1

        ttk.Label(settings_main, text="仕様3 (最高級) 割増率:").grid(row=row, column=0, sticky=tk.W, pady=8)
        ttk.Entry(settings_main, textvariable=self.spec3_rate, width=10).grid(row=row, column=1, sticky=tk.W, pady=8)
        ttk.Scale(settings_main, from_=0.20, to=0.80, orient=tk.HORIZONTAL, variable=self.spec3_rate, length=150).grid(
            row=row, column=2, sticky=tk.W, pady=8, padx=(10, 0))
        row += 1

        self.spec1_display = ttk.Label(settings_main, text="", font=("Arial", 9), foreground="blue")
        self.spec1_display.grid(row=row, column=1, sticky=tk.W, pady=2)
        row += 1
        self.spec2_display = ttk.Label(settings_main, text="", font=("Arial", 9), foreground="blue")
        self.spec2_display.grid(row=row, column=1, sticky=tk.W, pady=2)
        row += 1
        self.spec3_display = ttk.Label(settings_main, text="", font=("Arial", 9), foreground="blue")
        self.spec3_display.grid(row=row, column=1, sticky=tk.W, pady=2)
        row += 2

        # 経年減価補正率の追加
        ttk.Label(settings_main, text="経年減価補正率", font=sub_title_font).grid(row=row, column=0,
                                                                                  columnspan=3, pady=(20, 5),
                                                                                  sticky=tk.W)
        row += 1

        # 経年減価補正率の説明
        ttk.Label(settings_main, text="建物の経過年数に応じた補正率の参照表").grid(row=row, column=0,
                                                                                   columnspan=3, sticky=tk.W, pady=5)
        row += 1

        # スクロール可能なフレームを作成
        table_frame = ttk.Frame(settings_main)
        table_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)

        # スクロールバーを追加
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Treeviewで表を作成
        columns = ("year", "rate")
        self.depreciation_tree = ttk.Treeview(table_frame, columns=columns, show="headings",
                                              height=15, yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.depreciation_tree.yview)

        # 列の設定
        self.depreciation_tree.heading("year", text="経過年数")
        self.depreciation_tree.heading("rate", text="経年減点補正率")
        self.depreciation_tree.column("year", width=120, anchor=tk.CENTER)
        self.depreciation_tree.column("rate", width=150, anchor=tk.CENTER)

        # データの定義(1年～75年)
        depreciation_data = [
            (1, 0.9579), (2, 0.9309), (3, 0.9038), (4, 0.8803), (5, 0.8569),
            (6, 0.8335), (7, 0.8100), (8, 0.7866), (9, 0.7632), (10, 0.7397),
            (11, 0.7163), (12, 0.6929), (13, 0.6695), (14, 0.6460), (15, 0.6225),
            (16, 0.5992), (17, 0.5757), (18, 0.5523), (19, 0.5288), (20, 0.5054),
            (21, 0.4820), (22, 0.4585), (23, 0.4388), (24, 0.4189), (25, 0.3992),
            (26, 0.3794), (27, 0.3596), (28, 0.3398), (29, 0.3228), (30, 0.3059),
            (31, 0.2916), (32, 0.2774), (33, 0.2631), (34, 0.2488), (35, 0.2345),
            (36, 0.2294), (37, 0.2243), (38, 0.2191), (39, 0.2140), (40, 0.2089),
            (41, 0.2071), (42, 0.2053), (43, 0.2036), (44, 0.2018), (45, 0.2000),
            (46, 0.2000), (47, 0.2000), (48, 0.2000), (49, 0.2000), (50, 0.2000),
            (51, 0.2000), (52, 0.2000), (53, 0.2000), (54, 0.2000), (55, 0.2000),
            (56, 0.2000), (57, 0.2000), (58, 0.2000), (59, 0.2000), (60, 0.2000),
            (61, 0.2000), (62, 0.2000), (63, 0.2000), (64, 0.2000), (65, 0.2000),
            (66, 0.2000), (67, 0.2000), (68, 0.2000), (69, 0.2000), (70, 0.2000),
            (71, 0.2000), (72, 0.2000), (73, 0.2000), (74, 0.2000), (75, 0.2000)
        ]

        # データを挿入(1年～75年)
        for year, rate in depreciation_data:
            self.depreciation_tree.insert("", tk.END, values=(f"{year}年", f"{rate:.4f}"))

        self.depreciation_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        row += 1

        button_frame = ttk.Frame(settings_main)
        button_frame.grid(row=row, column=0, columnspan=3, pady=30)

        ttk.Button(button_frame, text="設定を初期値にリセット", command=self.reset_settings).pack(side=tk.LEFT,
                                                                                                  padx=(0, 10))
        ttk.Button(button_frame, text="設定内容を確認", command=self.show_current_settings).pack(side=tk.LEFT)
        row += 1

        self.misc_cost_rate.trace('w', self.update_settings_display)
        self.spec1_rate.trace('w', self.update_settings_display)
        self.spec2_rate.trace('w', self.update_settings_display)
        self.spec3_rate.trace('w', self.update_settings_display)

        self.update_settings_display()
    # 経費設定のタブに関する項目終了

    # LCC設定のタブに関する項目開始
    def create_lcc_tab(self):
        """LCC分析タブの作成"""
        canvas = tk.Canvas(self.lcc_frame)
        scrollbar = ttk.Scrollbar(self.lcc_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        row = 0



        # タイトル
        ttk.Label(main_frame, text="ライフサイクルコスト", font=title_font).grid(
            row=row, column=0, columnspan=3, pady=(0, 20), sticky=tk.W)
        row += 1

        # 説明
        desc_label = ttk.Label(main_frame, text="建物の建設費用を詳細に計算し、ライフサイクル全体のコストを分析します。",
                               font=sheet_font)
        desc_label.grid(row=row, column=0, columnspan=3, pady=(0, 20), sticky=tk.W)
        row += 1

        # ========== 基準単価と工事費率設定 ==========
        ttk.Label(main_frame, text="基準単価と修繕費用", font=sub_title_font).grid(
            row=row, column=0, columnspan=3, pady=(10, 10), sticky=tk.W)
        row += 1

        # ★★★ 統合フレームを作成 ★★★
        integrated_frame = ttk.Frame(main_frame)
        integrated_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))

        # 上段：基準単価フレーム（左側）
        basic_frame = ttk.Frame(integrated_frame)
        basic_frame.grid(row=0, column=0, sticky=(tk.W, tk.N), padx=(0, 20))

        # 再調達価格
        ttk.Label(basic_frame, text="再調達価格:", width=20).grid(row=0, column=0, sticky=tk.W, pady=3)
        ttk.Entry(basic_frame, textvariable=self.new_construction_cost, width=15, state='readonly',
                  justify=tk.RIGHT).grid(row=0, column=1, sticky=tk.W, padx=5, pady=3)
        ttk.Label(basic_frame, text="円").grid(row=0, column=2, sticky=tk.W, pady=3)

        # 建築基準単価
        ttk.Label(basic_frame, text="建築基準単価:", width=20).grid(row=1, column=0, sticky=tk.W, pady=3)
        ttk.Entry(basic_frame, textvariable=self.lcc_construction_unit_price, width=15, state='readonly',
                  justify=tk.RIGHT).grid(row=1, column=1, sticky=tk.W, padx=5, pady=3)
        ttk.Label(basic_frame, text="円/m2").grid(row=1, column=2, sticky=tk.W, pady=3)
        ttk.Label(basic_frame, textvariable=self.lcc_region_rate_display, font=sheet_font).grid(row=1, column=3,
                                                                                                sticky=tk.W,
                                                                                                padx=(10, 0), pady=3)

        # 地域掛け率表示を更新
        region_rate = self.get_region_rate()
        self.lcc_region_rate_display.set(f"建設地域係数: {region_rate:.2f}")

        # 建物延床面積
        ttk.Label(basic_frame, text="建物延床面積:", width=20).grid(row=2, column=0, sticky=tk.W, pady=3)
        ttk.Entry(basic_frame, textvariable=self.lcc_building_area, width=15, state='readonly', justify=tk.RIGHT).grid(
            row=2, column=1, sticky=tk.E, padx=5, pady=3)
        ttk.Label(basic_frame, text="m2").grid(row=2, column=2, sticky=tk.W, pady=3)

        # 仕様グレード
        ttk.Label(basic_frame, text="仕様グレード:", width=20).grid(row=3, column=0, sticky=tk.W, pady=3)
        ttk.Entry(basic_frame, textvariable=self.selected_spec_grade, width=15, state='readonly',
                  justify=tk.RIGHT).grid(row=3, column=1, sticky=tk.E, padx=5, pady=3)

        # 建物構造
        ttk.Label(basic_frame, text="建物構造:", width=18).grid(row=4, column=0, sticky=tk.W, pady=3)
        ttk.Entry(basic_frame, textvariable=self.structure_type, width=15, state='readonly', justify=tk.RIGHT).grid(
            row=4, column=1, sticky=tk.E, padx=5, pady=3)

        # ★★★ 円グラフ表示エリアを右側に配置（rowspan=2で2行分にまたがる） ★★★
        chart_frame = ttk.LabelFrame(integrated_frame, text="LCCコスト構成", padding="10")
        chart_frame.grid(row=0, column=1, rowspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(20, 0))

        # 円グラフ用のCanvasを作成
        self.lcc_pie_chart_canvas = tk.Canvas(chart_frame, width=300, height=300, bg='white')
        self.lcc_pie_chart_canvas.pack()

        # 凡例用フレーム
        legend_frame = ttk.Frame(chart_frame)
        legend_frame.pack(pady=(10, 0))

        # 凡例ラベル（初期値）
        self.lcc_legend_label1 = ttk.Label(legend_frame, text="■ 再調達価格: 0円", font=sheet_font)
        self.lcc_legend_label1.grid(row=0, column=0, sticky=tk.W, pady=2)
        self.lcc_legend_label2 = ttk.Label(legend_frame, text="■ 修繕工事費: 0円", font=sheet_font)
        self.lcc_legend_label2.grid(row=1, column=0, sticky=tk.W, pady=2)
        self.lcc_legend_label3 = ttk.Label(legend_frame, text="■ 運用費: 0円", font=sheet_font)
        self.lcc_legend_label3.grid(row=2, column=0, sticky=tk.W, pady=2)
        self.lcc_legend_label4 = ttk.Label(legend_frame, text="■ 一般管理費: 0円", font=sheet_font)
        self.lcc_legend_label4.grid(row=3, column=0, sticky=tk.W, pady=2)

        # 下段：再調達価格と各経費率フレーム
        rate_outer_frame = ttk.Frame(integrated_frame)
        rate_outer_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(15, 0))

        # セクションタイトル
        ttk.Label(rate_outer_frame, text="再調達価格と各経費率", font=sub_title_font).grid(
            row=0, column=0, columnspan=3, pady=(0, 10), sticky=tk.W)

        rate_frame = ttk.Frame(rate_outer_frame)
        rate_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 0))

        # 直接工事費
        ttk.Label(rate_frame, text="直接工事費:", width=20).grid(row=0, column=0, sticky=tk.W, pady=3)
        self.direct_construction_cost_rate = tk.StringVar(value="0.00")
        ttk.Entry(rate_frame, textvariable=self.direct_construction_cost_rate, width=15, state='readonly',
                  justify=tk.RIGHT).grid(
            row=0, column=1, sticky=tk.W, padx=5, pady=3)
        ttk.Label(rate_frame, text="%").grid(row=0, column=2, sticky=tk.W, pady=3)
        self.direct_construction_cost_amount = tk.StringVar(value="0")
        ttk.Entry(rate_frame, textvariable=self.direct_construction_cost_amount, width=15, state='readonly',
                  justify=tk.RIGHT).grid(
            row=0, column=3, sticky=tk.W, padx=5, pady=3)
        ttk.Label(rate_frame, text="円").grid(row=0, column=4, sticky=tk.W, pady=3)

        # 純工事費
        ttk.Label(rate_frame, text="純工事費:", width=20).grid(row=1, column=0, sticky=tk.W, pady=3)
        self.net_construction_cost_rate = tk.StringVar(value="80.37")
        ttk.Entry(rate_frame, textvariable=self.net_construction_cost_rate, width=15, state='readonly',
                  justify=tk.RIGHT).grid(
            row=1, column=1, sticky=tk.W, padx=5, pady=3)
        ttk.Label(rate_frame, text="%").grid(row=1, column=2, sticky=tk.W, pady=3)
        self.net_construction_cost_amount = tk.StringVar(value="0")
        ttk.Entry(rate_frame, textvariable=self.net_construction_cost_amount, width=15, state='readonly',
                  justify=tk.RIGHT).grid(
            row=1, column=3, sticky=tk.W, padx=5, pady=3)
        ttk.Label(rate_frame, text="円").grid(row=1, column=4, sticky=tk.W, pady=3)

        # 現場管理費
        ttk.Label(rate_frame, text="現場管理費:", width=20).grid(row=2, column=0, sticky=tk.W, pady=3)
        self.site_management_cost_rate = tk.StringVar(value="7")
        ttk.Entry(rate_frame, textvariable=self.site_management_cost_rate, width=15, justify=tk.RIGHT).grid(
            row=2, column=1, sticky=tk.W, padx=5, pady=3)
        ttk.Label(rate_frame, text="%").grid(row=2, column=2, sticky=tk.W, pady=3)
        self.site_management_cost_amount = tk.StringVar(value="0")
        ttk.Entry(rate_frame, textvariable=self.site_management_cost_amount, width=15, state='readonly',
                  justify=tk.RIGHT).grid(
            row=2, column=3, sticky=tk.W, padx=5, pady=3)
        ttk.Label(rate_frame, text="円").grid(row=2, column=4, sticky=tk.W, pady=3)

        # 共通仮設費
        ttk.Label(rate_frame, text="共通仮設費:", width=20).grid(row=3, column=0, sticky=tk.W, pady=3)
        self.common_temporary_cost_rate = tk.StringVar(value="4")
        ttk.Entry(rate_frame, textvariable=self.common_temporary_cost_rate, width=15, justify=tk.RIGHT).grid(
            row=3, column=1, sticky=tk.W, padx=5, pady=3)
        ttk.Label(rate_frame, text="%").grid(row=3, column=2, sticky=tk.W, pady=3)
        self.common_temporary_cost_amount = tk.StringVar(value="0")
        ttk.Entry(rate_frame, textvariable=self.common_temporary_cost_amount, width=15, state='readonly',
                  justify=tk.RIGHT).grid(
            row=3, column=3, sticky=tk.W, padx=5, pady=3)
        ttk.Label(rate_frame, text="円").grid(row=3, column=4, sticky=tk.W, pady=3)

        # 一般管理費等
        ttk.Label(rate_frame, text="一般管理費等:", width=20).grid(row=4, column=0, sticky=tk.W, pady=3)
        self.general_admin_fee_rate = tk.StringVar(value="12")
        ttk.Entry(rate_frame, textvariable=self.general_admin_fee_rate, width=15, justify=tk.RIGHT).grid(
            row=4, column=1, sticky=tk.W, padx=5, pady=3)
        ttk.Label(rate_frame, text="%").grid(row=4, column=2, sticky=tk.W, pady=3)
        self.general_admin_fee_amount = tk.StringVar(value="0")
        ttk.Entry(rate_frame, textvariable=self.general_admin_fee_amount, width=15, state='readonly',
                  justify=tk.RIGHT).grid(
            row=4, column=3, sticky=tk.W, padx=5, pady=3)
        ttk.Label(rate_frame, text="円").grid(row=4, column=4, sticky=tk.W, pady=3)

        row += 1

        # ========== LCC修繕計画表システム（整理版） ==========

        # ========== 10行19列 LCC修繕計画表 ==========
        ttk.Label(main_frame, text="LCC修繕計画表", font=sub_title_font).grid(
            row=row, column=0, columnspan=3, pady=(15, 10), sticky=tk.W)
        row += 1

        # 表全体を保持するフレームを作成
        table_frame = ttk.Frame(main_frame)
        table_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 15))

        # ヘッダーと工種名の定義
        HEADER_TEXTS = [
            "工種", "比率", "項目別再調達価格",
            "0-5", "6-10", "11-15", "16-20", "21-25", "26-30",
            "31-35", "36-40", "41-45", "46-50", "51-55", "56-60",
            "61-65", "66-70", "71-75", "計"
        ]

        COMPONENT_NAMES = [
            "地業工事", "躯体工事", "防水工事", "外装工事", "内装工事", "建具工事", "外構工事", "直接仮設費",
            "建築工事計",
            "熱源機器設備", "熱源配管設備", "空調機器設備", "空調配管設備", "換気機器設備", "換気ダクト設備",
            "自動制御設備", "排煙設備", "空調設備工事計",
            "衛生器具設備", "給水機器設備", "給水配管設備", "給湯機器設備", "給湯配管設備", "排水設備", "ガス設備",
            "厨房設備", "浄化槽設備", "消火設備", "油送設備", "衛生設備計",
            "受変電設備", "自家発電設備", "動力設備", "電灯設備", "情報通信設備", "中央監視制御設備", "火災報知設備",
            "電気設備計",
            "エレベータ設備", "ダムウェーダー設備", "昇降機設備計", "直接工事計",
            "運用費",
            "一般管理費"
        ]

        # ★★★ 【追加】列幅の定義（グローバル変数） ★★★
        COLUMN_WIDTHS = {
            0: 10,  # 工種列
            1: 5,  # 比率列
            2: 12,  # 項目別再調達価格列（幅を広げた）
            3: 5,  # 年度期間列（0-5年）
            4: 5,  # 年度期間列（6-10年）
            5: 5,  # 年度期間列（11-15年）
            6: 5,  # 年度期間列（16-20年）
            7: 5,  # 年度期間列（21-25年）
            8: 5,  # 年度期間列（26-30年）
            9: 5,  # 年度期間列（31-35年）
            10: 5,  # 年度期間列（36-40年）
            11: 5,  # 年度期間列（41-45年）
            12: 5,  # 年度期間列（46-50年）
            13: 5,  # 年度期間列（51-55年）
            14: 5,  # 年度期間列（56-60年）
            15: 5,  # 年度期間列（61-65年）
            16: 5,  # 年度期間列（66-70年）
            17: 5,  # 年度期間列（71-75年）
            18: 8  # 計列
        }

        # 維持管理費用対象工種の定義
        # 行番号でマッピング（1-indexed、0行目はヘッダー）
        MAINTENANCE_TARGET_ROWS = {
            11: "熱源配管設備",  # 11行目
            13: "空調配管設備",  # 13行目
            21: "給水配管設備",  # 21行目
            23: "給湯配管設備",  # 23行目
            24: "排水設備",  # 24行目
            34: "電灯設備"  # 34行目
        }

        # ★★★ 【追加】大規模修繕費用対象工種の定義 ★★★
        MAJOR_REPAIR_TARGET_ROWS = {
            3: "防水工事",
            4: "外装工事",
            10: "熱源機器設備",
            11: "熱源配管設備",
            12: "空調機器設備",
            13: "空調配管設備",
            14: "換気機器設備",
            16: "自動制御設備",
            19: "衛生器具設備",
            20: "給水機器設備",
            21: "給水配管設備",
            22: "給湯機器設備",
            23: "給湯配管設備",
            24: "排水設備",
            27: "浄化槽設備",
            28: "消火設備",
            29: "油送設備",
            31: "受変電設備",
            32: "自家発電設備",
            33: "動力設備",
            34: "電灯設備",
            35: "情報通信設備",
            36: "中央監視制御設備",
            37: "火災報知設備",
            39: "エレベータ設備",
            40: "ダムウェーダー設備"
        }

        # ★★★ 【追加】統合対象工種（維持管理 + 大規模修繕）★★★
        # 両方の対象となる工種をマージ
        COMBINED_TARGET_ROWS = {}
        COMBINED_TARGET_ROWS.update(MAINTENANCE_TARGET_ROWS)
        COMBINED_TARGET_ROWS.update(MAJOR_REPAIR_TARGET_ROWS)
        # ★★★ 【追加終了】 ★★★

        # 表の構造定義（各工事区分の範囲を定義）
        WORK_SECTIONS = {
            'architectural': {'start': 1, 'end': 8, 'total_row': 9, 'color': 'pink'},
            'hvac': {'start': 10, 'end': 17, 'total_row': 18, 'color': 'lightblue'},
            'sanitary': {'start': 19, 'end': 29, 'total_row': 30, 'color': 'lightgreen'},
            'electrical': {'start': 31, 'end': 37, 'total_row': 38, 'color': 'lightyellow'},
            'elevator': {'start': 39, 'end': 40, 'total_row': 41, 'color': 'plum'},
            'grand_total': {'row': 42, 'color': 'lightgray'},
            'operation_cost': {'row': 43, 'color': 'lightcyan'},
            'general_admin_cost': {'row': 44, 'color': 'lightyellow'}
        }

        # 比率列の初期値を定義
        DEFAULT_RATIOS = {
            # 建築工事（新築工事費の40%）
            1: 9.0, 2: 37.0, 3: 9.0, 4: 20.0, 5: 11.0, 6: 5.0, 7: 3.0, 8: 6.0,
            9: 40.0,  # 建築工事計★
            # 空調設備工事（新築工事費の12%）
            10: 15.0, 11: 10.0, 12: 25.0, 13: 13.0, 14: 8.0, 15: 6.0, 16: 15.0, 17: 8.0,
            18: 12.0,  # 空調設備工事計★
            # 衛生設備工事（新築工事費の8%）
            19: 8.0, 20: 8.0, 21: 8.0, 22: 11.0, 23: 8.0, 24: 19.0, 25: 6.0, 26: 5.0, 27: 0, 28: 25.0, 29: 2.0,
            30: 8.0,  # 衛生設備計★
            # 電気設備工事（新築工事費の15.5%）
            31: 10.0, 32: 8.0, 33: 20.0, 34: 20.0, 35: 15.0, 36: 12.0, 37: 15.0,
            38: 15.5,  # 電気設備計★
            # 昇降機設備工事（新築工事費の0.87%）
            39: 80.0, 40: 20.0,
            41: 0.87,  # 昇降機設備計★
            42: 76.37,  # 直接工事計
            # 43: 運用費（比率なし）
        }

        # 現在の比率を保存する変数
        self.current_ratios = DEFAULT_RATIOS.copy()

        # 経過年数の範囲定義を追加
        self.year_ranges = {
            3: (0, 5),
            4: (6, 10),
            5: (11, 15),
            6: (16, 20),
            7: (21, 25),
            8: (26, 30),
            9: (31, 35),
            10: (36, 40),
            11: (41, 45),
            12: (46, 50),
            13: (51, 55),
            14: (56, 60),
            15: (61, 65),
            16: (66, 70),
            17: (71, 75)
        }

        NUM_ROWS = 45
        NUM_COLS = 19

        # 経年減価補正率表（経費設定タブと同じデータ）
        self.depreciation_rates = {}
        for year in range(1, 45):
            # 1年～44年のデータ（元のデータ）
            depreciation_data = {
                1: 0.9579, 2: 0.9309, 3: 0.9038, 4: 0.8803, 5: 0.8569,
                6: 0.8335, 7: 0.8100, 8: 0.7866, 9: 0.7632, 10: 0.7397,
                11: 0.7163, 12: 0.6929, 13: 0.6695, 14: 0.6460, 15: 0.6225,
                16: 0.5992, 17: 0.5757, 18: 0.5523, 19: 0.5288, 20: 0.5054,
                21: 0.4820, 22: 0.4585, 23: 0.4388, 24: 0.4189, 25: 0.3992,
                26: 0.3794, 27: 0.3596, 28: 0.3398, 29: 0.3228, 30: 0.3059,
                31: 0.2916, 32: 0.2774, 33: 0.2631, 34: 0.2488, 35: 0.2345,
                36: 0.2294, 37: 0.2243, 38: 0.2191, 39: 0.2140, 40: 0.2089,
                41: 0.2071, 42: 0.2053, 43: 0.2036, 44: 0.2018
            }
            self.depreciation_rates[year] = depreciation_data.get(year, 0.2000)

        # 45年以上は0.2000
        for year in range(45, 76):
            self.depreciation_rates[year] = 0.2000

        # ========== 【修正5】一般管理費計算関数の追加 ==========
        def calculate_general_admin_cost(self, operation_cost_monthly=None):
            """
            一般管理費を計算して表に反映
            """
            try:
                import datetime


                # 再調達価格を取得
                reconstruction_cost_str = self.new_construction_cost.get().replace(',', '').strip()
                if not reconstruction_cost_str:
                    reconstruction_cost = 0.0
                else:
                    reconstruction_cost = float(reconstruction_cost_str)

                # 固定資産税率を取得
                tax_rate_str = self.lcc_tax_rate.get().replace(',', '').strip()
                if not tax_rate_str:
                    tax_rate = 0.0
                else:
                    tax_rate = float(tax_rate_str) / 100

                # 保険料率を取得
                insurance_rate_str = self.lcc_insurance_rate.get().replace(',', '').strip()
                if not insurance_rate_str:
                    insurance_rate = 0.0
                else:
                    insurance_rate = float(insurance_rate_str) / 100

                # ========== 【修正】運営費合計を引数から取得、なければ変数から取得 ==========
                if operation_cost_monthly is None:
                    operation_cost_monthly_str = self.lcc_operation_cost.get().replace(',', '').strip()
                    if not operation_cost_monthly_str:
                        operation_cost_monthly = 0.0
                    else:
                        operation_cost_monthly = float(operation_cost_monthly_str)

                # 5年分（60ヶ月分）の運営費を100万円単位に変換
                operation_cost_per_period = operation_cost_monthly * 60
                operation_cost_million = operation_cost_per_period / 1_000_000


                print(f"再調達価格: {reconstruction_cost:,.0f}円")
                print(f"固定資産税率: {tax_rate * 100}%")
                print(f"保険料率: {insurance_rate * 100}%")
                print(f"運営費（月額）: {operation_cost_monthly:,.0f}円")
                print(f"運営費（5年分）: {operation_cost_million:.2f}百万円")

                # 44行目（一般管理費行）
                admin_row = 44

                # 1列目と2列目に"-"を設定
                self.table_vars[admin_row][1].set("-")
                self.table_vars[admin_row][2].set("-")

                # 3列目～17列目（年度期間）に一般管理費を設定
                period_total = 0

                for col in range(3, 18):
                    # 年度範囲を取得（例: 3列目 = (0, 5)年）
                    year_range = self.year_ranges.get(col, (0, 0))
                    start_year, end_year = year_range

                    # 期間の中央年を使用
                    mid_year = (start_year + end_year) // 2

                    # 経年減価補正率を取得
                    depreciation_rate = self.depreciation_rates.get(mid_year, 0.2000)

                    # ========== 【修正】運用費は引数で受け取った運営費を使用 ==========
                    # operation_cost_million は既にforループの前で計算済み
                    # （operation_cost_monthly * 60 / 1_000_000）
                    operation_cost_for_period = operation_cost_million
                    # ========== 修正ここまで ==========

                    # 固定資産税（5年分）
                    property_tax_per_period = reconstruction_cost * depreciation_rate * tax_rate * 5
                    property_tax_million = property_tax_per_period / 1_000_000

                    # 保険料（5年分）
                    insurance_per_period = reconstruction_cost * depreciation_rate * insurance_rate * 5
                    insurance_million = insurance_per_period / 1_000_000

                    # 一般管理費 = 運用費 + 固定資産税 + 保険料
                    admin_cost_million = operation_cost_for_period + property_tax_million + insurance_million

                    self.table_vars[admin_row][col].set(f"{admin_cost_million:.2f}")
                    period_total += admin_cost_million

                    print(f"列{col} ({start_year}-{end_year}年, 中央{mid_year}年):")
                    print(f"  経年減価補正率: {depreciation_rate:.4f}")
                    print(f"  運用費（運営費）: {operation_cost_for_period:.2f}百万円")
                    print(f"  固定資産税: {property_tax_million:.2f}百万円")
                    print(f"  保険料: {insurance_million:.2f}百万円")
                    print(f"  一般管理費合計: {admin_cost_million:.2f}百万円")
                    print(f"  固定資産税: 再調達価格 × 経年減価補正率 × 固定資産税率 × 5年")
                    print(f"  保険料: 再調達価格 × 経年減価補正率 × 保険料率 × 5年")

                # 18列目（計列）に合計を設定
                self.table_vars[admin_row][18].set(f"{period_total:.2f}")
                self.lcc_frame.update_idletasks()

            except Exception as e:
                print(f"一般管理費計算エラー: {e}")
                import traceback
                traceback.print_exc()





        # ↓↓↓ ステップ1: 43行19列分の tk.StringVar を定義（これが最初！） ↓↓↓
        self.table_vars = []
        for r in range(NUM_ROWS):
            row_vars = []
            for c in range(NUM_COLS):
                initial_value = ""
                if r == 0:
                    initial_value = HEADER_TEXTS[c]
                elif c == 0 and r - 1 < len(COMPONENT_NAMES):
                    initial_value = COMPONENT_NAMES[r - 1]

                row_vars.append(tk.StringVar(value=initial_value))
            self.table_vars.append(row_vars)
        # ↑↑↑ StringVarの定義完了 ↑↑↑

        # ↓↓↓ ステップ2: 経過年数マーカー表示用のラベルを保持する辞書を初期化 ↓↓↓
        self.year_marker_labels = {}

        # ↓↓↓ ステップ3: マーカー行（table_frameの0行目）を作成 ★★★ 注釈も同時に表示 ★★★ ↓↓↓
        for c in range(NUM_COLS):
            # グローバル変数から幅を取得
            entry_width = COLUMN_WIDTHS.get(c, 5)

            # ★★★ 18列目（計列）には「100万円単位」と表示 ★★★
            if c == 18:
                marker_label = tk.Label(
                    table_frame,
                    text="100万円単位",
                    width=entry_width,
                    justify=tk.CENTER,
                    fg='black',
                    font=sheet_font,
                    relief=tk.FLAT
                )
            else:
                # その他の列はマーカー用（空白、▼が入る）
                marker_label = tk.Label(
                    table_frame,
                    text="",
                    width=entry_width,
                    justify=tk.CENTER,
                    fg='red',
                    font=("Arial", 12, "bold"),
                    relief=tk.FLAT
                )

            marker_label.grid(row=0, column=c, padx=0, pady=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            self.year_marker_labels[c] = marker_label
        # ↑↑↑ マーカー行の作成完了 ↑↑↑

        # ↓↓↓ ステップ4: 既存の表（ヘッダー+データ）を1行目以降に配置 ↓↓↓
        for r in range(NUM_ROWS):
            for c in range(NUM_COLS):
                # グローバル変数から幅を取得
                entry_width = COLUMN_WIDTHS.get(c, 5)  # デフォルトは5
                # またはリスト形式の場合: entry_width = COLUMN_WIDTHS[c]

                # ★★★ 【変更】読み取り専用の判定ロジック ★★★
                # 読み込み専用とする行番号のリスト
                readonly_rows = [
                    9,  # 建築工事計
                    18,  # 空調設備工事計
                    30,  # 衛生設備計
                    38,  # 電気設備計
                    41,  # 昇降機設備計
                    42,  # 直接工事計
                    43,  # 運用費
                    44   # 一般管理費
                ]

                if r == 0:
                    # ヘッダー行は全て読み取り専用
                    is_readonly = True
                elif r == 43:  # 運用費行
                    # 運用費行は全列読み取り専用
                    is_readonly = True
                elif c == 0:
                    # 工種名列（0列目）は全て読み取り専用
                    is_readonly = True
                elif c == 2:
                    # ★★★ 【追加】項目別再調達価格列（2列目）は全て読み取り専用 ★★★
                    is_readonly = True
                elif r in readonly_rows:
                    # ★★★ 【追加】各工事計行は全列読み取り専用 ★★★
                    is_readonly = True
                elif r in COMBINED_TARGET_ROWS:
                    # 統合対象工種（維持管理 + 大規模修繕）
                    if c == 1:
                        # 比率列は編集可能
                        is_readonly = False
                    elif 3 <= c <= 18:
                        # 年度期間列と計列は読み取り専用（自動計算）
                        is_readonly = True
                    else:
                        is_readonly = True
                elif c == 1 and r > 0:
                    # 比率列の読み取り専用判定
                    # 各工事計行は既にreadonly_rowsで処理済み
                    is_readonly = False
                elif 3 <= c <= 18 and r > 0:
                    # 年度期間列（3-17列）と計列（18列）
                    # 統合対象行以外は編集可能（ただし統合対象は上で処理済み）
                    if r in COMBINED_TARGET_ROWS:
                        is_readonly = True
                    else:
                        is_readonly = False
                else:
                    # その他のセル
                    is_readonly = True
                # ★★★ 【変更終了】 ★★★

                # 背景色を決定
                background_color = 'white'
                read_only_background_color = 'white'

                for section_name, section_info in WORK_SECTIONS.items():
                    if section_name == 'grand_total':
                        if r == section_info['row']:
                            background_color = section_info['color']
                            read_only_background_color = section_info['color']
                    elif section_name == 'operation_cost':
                        if r == section_info['row']:
                            background_color = section_info['color']
                            read_only_background_color = section_info['color']
                        # ========== 【追加】一般管理費行の背景色設定 ==========
                    elif section_name == 'general_admin_cost':
                        if r == section_info['row']:
                            background_color = section_info['color']
                            read_only_background_color = section_info['color']
                        # ========== 追加ここまで ==========
                    else:
                        # 'start'と'end'を持つセクションのみ処理
                        if 'start' in section_info and 'end' in section_info:
                            if c == 0 and section_info['start'] <= r <= section_info['end']:
                                background_color = section_info['color']
                                read_only_background_color = section_info['color']
                            if r == section_info['total_row']:
                                background_color = section_info['color']
                                read_only_background_color = section_info['color']

                entry = tk.Entry(
                    table_frame,
                    textvariable=self.table_vars[r][c],
                    width=entry_width,  # ← グローバル変数を使用
                    state='readonly' if is_readonly else 'normal',
                    justify=tk.CENTER if r == 0 else tk.RIGHT,
                    bg=background_color,
                    readonlybackground=read_only_background_color,
                    fg='black',
                    borderwidth=1,
                    relief=tk.RIDGE
                )

                entry.grid(row=r + 2, column=c, padx=0, pady=0, sticky=(tk.W, tk.E, tk.N, tk.S))

                # 比率列（1列目）の変更監視
                if c == 1 and r > 0 and not is_readonly:
                    self.table_vars[r][c].trace_add('write', lambda *args, row=r: on_ratio_change(self, row))

                # 年度期間列（3-17列）の変更監視
                # 統合対象行と各工事計行は編集不可なので監視不要
                if 3 <= c <= 17 and r > 0 and not is_readonly and r not in COMBINED_TARGET_ROWS and r not in readonly_rows:
                    self.table_vars[r][c].trace_add('write',
                                                    lambda *args, row=r, col=c: on_period_change(self, row, col))
                # ★★★ 【変更終了】 ★★★
        row += 1

        # ========== 比率列の操作ボタン ==========
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 10))

        reset_ratio_btn = ttk.Button(button_frame, text="比率を初期値に戻す", command=lambda: reset_ratios(self))
        reset_ratio_btn.grid(row=0, column=0, padx=5)

        save_ratio_btn = ttk.Button(button_frame, text="現在の比率を保存", command=lambda: save_ratios(self))
        save_ratio_btn.grid(row=0, column=1, padx=5)

        load_ratio_btn = ttk.Button(button_frame, text="保存した比率を読み込む", command=lambda: load_ratios(self))
        load_ratio_btn.grid(row=0, column=2, padx=5)

        row += 1

        # ========== 経過年数マーカー更新関数 ==========
        def update_lcc_year_marker(self, elapsed_years):
            """
            経過年数に応じて、該当する列の上に▼マーカーを表示
            elapsed_years: 経過年数（float）
            """
            try:
                # まず全てのマーカーをクリア（18列目以外）
                if not hasattr(self, 'year_marker_labels'):
                    return

                for col_idx, label in self.year_marker_labels.items():
                    if col_idx == 18:
                        # 18列目は「100万円単位」を保持
                        label.config(text="100万円単位")
                    else:
                        # その他の列はクリア
                        label.config(text="")

                # 経過年数が0以下の場合は何も表示しない
                if elapsed_years <= 0:
                    return

                elapsed_years_int = int(elapsed_years)  # 切り下げ
                print(f"経過年数: {elapsed_years:.1f}年 → 切り下げ: {elapsed_years_int}年")

                # 該当する列を検索
                target_col = None
                for col_idx, (start_year, end_year) in self.year_ranges.items():
                    if start_year <= elapsed_years_int <= end_year:
                        target_col = col_idx
                        break

                # 該当列が見つかった場合、▼を表示
                if target_col is not None:
                    self.year_marker_labels[target_col].config(text="▼")
                    # ↓↓↓ メッセージも修正 ↓↓↓
                    print(f"経過年数 {elapsed_years_int}年 → 列{target_col}に▼を表示")
                    # ↑↑↑ 修正 ↑↑↑
                else:
                    # ↓↓↓ メッセージも修正 ↓↓↓
                    print(f"経過年数 {elapsed_years_int}年は表示範囲外です（0-75年）")
                    # ↑↑↑ 修正 ↑↑↑

            except Exception as e:
                print(f"経過年数マーカー更新エラー: {e}")
                import traceback
                traceback.print_exc()

        # ========== 運用費計算関数 ==========
        def calculate_operation_cost(self):
            """
            運用費を計算して表に反映
            運用費 = (エネルギー単価 + 水道単価) × 延床面積 × 12ヶ月 × 5年
            100万円単位で表示
            """
            try:
                # 延床面積を取得
                area_str = self.lcc_building_area.get().replace(',', '').strip()
                if not area_str:
                    area = 0.0
                else:
                    area = float(area_str)

                # エネルギー単価を取得
                energy_str = self.lcc_energy_unit_cost.get().replace(',', '').strip()
                if not energy_str:
                    energy_unit = 0.0
                else:
                    energy_unit = float(energy_str)

                # 水道単価を取得
                water_str = self.lcc_water_unit_cost.get().replace(',', '').strip()
                if not water_str:
                    water_unit = 0.0
                else:
                    water_unit = float(water_str)

                # 運用費の計算
                # (エネルギー単価 + 水道単価) × 延床面積 × 12ヶ月 × 5年
                operation_cost_per_period = (energy_unit + water_unit) * area * 12 * 5

                # 100万円単位に変換
                operation_cost_million = operation_cost_per_period / 1_000_000

                # 43行目（運用費行）に設定
                operation_row = 43

                # 1列目と2列目に"-"を設定
                self.table_vars[operation_row][1].set("-")
                self.table_vars[operation_row][2].set("-")

                # 3列目～17列目（年度期間）に運用費を設定
                period_total = 0
                for col in range(3, 18):
                    self.table_vars[operation_row][col].set(f"{operation_cost_million:.2f}")
                    period_total += operation_cost_million

                # 18列目（計列）に合計を設定
                self.table_vars[operation_row][18].set(f"{period_total:.2f}")

                print(f"運用費計算完了: {operation_cost_million:.2f}百万円/期間")


            except Exception as e:
                print(f"運用費計算エラー: {e}")
                import traceback
                traceback.print_exc()

        # ========== 運営費合計計算関数 ==========
        def calculate_total_operation_cost(self):
            """
            運営費の各内訳を合計して、運営費（合計）に反映
            """
            try:

                # 各内訳を取得
                personnel = self._parse_number(self.lcc_personnel_cost.get())
                amenity = self._parse_number(self.lcc_amenity_cost.get())
                ota_ad = self._parse_number(self.lcc_ota_ad_cost.get())
                food = self._parse_number(self.lcc_food_cost.get())
                linen = self._parse_number(self.lcc_linen_cost.get())

                # 合計を計算
                total = personnel + amenity + ota_ad + food + linen

                # 運営費（合計）に設定（カンマ区切り）
                self.lcc_operation_cost.set(f"{total:,.0f}")

                print(f"運営費合計更新: {total:,.0f}円/月")
                self.calculate_general_admin_cost(operation_cost_monthly=total)

            except Exception as e:
                print(f"運営費合計計算エラー: {e}")
                import traceback
                traceback.print_exc()

        # ★★★ 【変更】維持管理費用と大規模修繕費用を統合した計算関数（修正版）★★★
        def calculate_maintenance_and_repair_costs(self):
            """
            維持管理費用と大規模修繕費用を統合して計算
            セルの値 = 年間維持管理費（5年分） + 大規模修繕費（該当がある場合）
            100万円単位で表示
            計算後、各工事区分の年度期間集計も実行
            """
            try:
                # print("\n=== calculate_maintenance_and_repair_costs 開始 ===")

                # 年間維持管理比率を取得
                maintenance_rate_str = self.lcc_annual_maintenance_rate.get().replace(',', '').strip()
                if not maintenance_rate_str:
                    maintenance_rate = 0.0
                else:
                    maintenance_rate = float(maintenance_rate_str) / 100  # %を小数に変換

                # 大規模修繕費率を取得
                repair_rate_str = self.lcc_major_repair_rate.get().replace(',', '').strip()
                if not repair_rate_str:
                    repair_rate = 0.0
                else:
                    repair_rate = float(repair_rate_str) / 100  # %を小数に変換

                # 大規模修繕周期を取得
                repair_interval_str = self.lcc_major_repair_interval.get().replace(',', '').strip()
                if not repair_interval_str or float(repair_interval_str) == 0:
                    repair_interval = 0
                else:
                    repair_interval = int(float(repair_interval_str))

                # print(f"年間維持管理比率: {maintenance_rate * 100}%")
                # print(f"大規模修繕費率: {repair_rate * 100}%")
                # print(f"大規模修繕周期: {repair_interval}年")

                # ★★★ 大規模修繕が発生する列を計算 ★★★
                repair_columns = []
                if repair_interval > 0:
                    # 基準年0年から開始
                    for col_idx, (start_year, end_year) in self.year_ranges.items():
                        # 周期の倍数が範囲内に含まれるかチェック
                        multiple = repair_interval
                        while multiple <= end_year:
                            if start_year <= multiple <= end_year:
                                if col_idx not in repair_columns:
                                    repair_columns.append(col_idx)
                                # print(f"  {multiple}年 → 列{col_idx} ({start_year}-{end_year}年)")
                                break  # この列には1回だけカウント
                            multiple += repair_interval

                # print(f"大規模修繕が発生する列: {repair_columns}")

                # 統合対象の各工種について計算
                for row_num, component_name in COMBINED_TARGET_ROWS.items():
                    # 2列目（金額列）の値を取得
                    amount_str = self.table_vars[row_num][2].get().replace(',', '').strip()
                    if not amount_str or amount_str == "":
                        amount = 0.0
                    else:
                        amount = float(amount_str)

                    # 維持管理費用対象かチェック
                    is_maintenance_target = row_num in MAINTENANCE_TARGET_ROWS
                    # 大規模修繕対象かチェック
                    is_repair_target = row_num in MAJOR_REPAIR_TARGET_ROWS

                    # print(f"\n行{row_num} {component_name}:")
                    # print(f"  金額: {amount:,.0f}円")
                    # print(f"  維持管理対象: {is_maintenance_target}")
                    # print(f"  大規模修繕対象: {is_repair_target}")

                    # 3列目～17列目（年度期間）に費用を設定
                    period_total = 0
                    for col in range(3, 18):
                        cost_million = 0.0

                        # 維持管理費用を加算（対象の場合）
                        if is_maintenance_target:
                            maintenance_cost_per_period = amount * maintenance_rate * 5
                            maintenance_cost_million = maintenance_cost_per_period / 1_000_000
                            cost_million += maintenance_cost_million

                        # 大規模修繕費用を加算（対象かつ該当列の場合）
                        if is_repair_target and col in repair_columns:
                            repair_cost = amount * repair_rate
                            repair_cost_million = repair_cost / 1_000_000
                            cost_million += repair_cost_million
                            # print(
                            #     f"  列{col}: 維持管理={maintenance_cost_million if is_maintenance_target else 0:.2f} + 大規模修繕={repair_cost_million:.2f} = {cost_million:.2f}百万円")

                        self.table_vars[row_num][col].set(f"{cost_million:.2f}" if cost_million > 0 else "")
                        period_total += cost_million

                    # 18列目（計列）に合計を設定
                    self.table_vars[row_num][18].set(f"{period_total:.2f}" if period_total > 0 else "")
                    # print(f"  合計: {period_total:.2f}百万円")

                # print("=== calculate_maintenance_and_repair_costs 完了 ===\n")

                self.calculate_section_period_totals()# 各工事区分と直接工事計の年度期間集計を実行


            except Exception as e:
                print(f"維持管理・大規模修繕費用計算エラー: {e}")
                import traceback
                traceback.print_exc()


        # ========== 比率変更時のコールバック関数 ==========
        def on_ratio_change(self, changed_row):
            """
            比率列が変更されたときに呼ばれる関数
            該当する工種の金額を再計算し、所属する工事区分の合計も更新する
            """
            try:
                # 新築工事推定金額を取得
                base_cost = self._parse_number(self.new_construction_cost.get())
                if base_cost <= 0:
                    return

                # 変更された行がどの工事区分に属するかを判定
                section_name = None
                for sec_name, sec_info in WORK_SECTIONS.items():
                    # ========== 【修正】スキップ条件を追加 ==========
                    if sec_name in ['grand_total', 'operation_cost', 'general_admin_cost']:
                        continue
                    # ========== 修正ここまで ==========

                    if sec_info['start'] <= changed_row <= sec_info['end']:
                        section_name = sec_name
                        break

                if not section_name:
                    return

                section_info = WORK_SECTIONS[section_name]
                total_row = section_info['total_row']

                # 工事区分の「計」の金額を取得
                section_total_ratio = self._parse_number(self.table_vars[total_row][1].get())
                section_total_amount = base_cost * (section_total_ratio / 100)

                # 変更された工種の金額を計算
                changed_ratio = self._parse_number(self.table_vars[changed_row][1].get())
                changed_amount = section_total_amount * (changed_ratio / 100)
                self.table_vars[changed_row][2].set(self._format_number(changed_amount))

                # 同じ工事区分内の全工種の金額を合計して「計」を更新
                start_row = section_info['start']
                end_row = section_info['end']
                section_sum = 0
                for r in range(start_row, end_row + 1):
                    amount = self._parse_number(self.table_vars[r][2].get())
                    section_sum += amount

                # 「計」行の金額を更新
                self.table_vars[total_row][2].set(self._format_number(section_sum))

                # 直接工事計を更新
                grand_total_row = WORK_SECTIONS['grand_total']['row']
                grand_sum = 0
                for sec_name, sec_info in WORK_SECTIONS.items():
                    # ========== 【修正】スキップ条件を追加 ==========
                    if sec_name in ['grand_total', 'operation_cost', 'general_admin_cost']:
                        continue
                    # ========== 修正ここまで ==========

                    total_row = sec_info['total_row']
                    amount = self._parse_number(self.table_vars[total_row][2].get())
                    grand_sum += amount

                self.table_vars[grand_total_row][2].set(self._format_number(grand_sum))

                # 直接工事計の比率も更新
                if base_cost > 0:
                    grand_ratio = (grand_sum / base_cost) * 100
                    self.table_vars[grand_total_row][1].set(f"{grand_ratio:.2f}%")

                # 純工事費率も更新
                self.update_net_construction_rate()

                # 統合対象行の場合は費用も再計算
                if changed_row in COMBINED_TARGET_ROWS:
                    self.calculate_maintenance_and_repair_costs()

            except Exception as e:
                print(f"比率変更エラー: {e}")
        # ★★★ 追加: 年度期間変更時のコールバック関数 ★★★
        # ★★★ 年度期間変更時のコールバック関数（修正版）★★★
        def on_period_change(self, changed_row, changed_col):
            """
            年度期間列（3～17列）が変更されたときに呼ばれる関数
            1. 該当工種の年度期間の合計を「計」列（18列目）に反映
            2. 同じ工事区分の「計」行の該当年度期間列を更新
            3. 「計」行の年度期間合計を「計」列に反映
            4. 直接工事計の該当年度期間列を更新
            5. 直接工事計の年度期間合計を「計」列に反映
            """
            try:
                # 変更された行がどの工事区分に属するかを判定
                section_name = None
                for sec_name, sec_info in WORK_SECTIONS.items():
                    if sec_name in ['grand_total', 'operation_cost', 'general_admin_cost']:
                        continue
                    if sec_info['start'] <= changed_row <= sec_info['end']:
                        section_name = sec_name
                        break

                if not section_name:
                    return

                section_info = WORK_SECTIONS[section_name]

                # 1. 該当工種の年度期間合計を計算（3～17列の合計 → 18列目）
                period_sum = 0
                for col in range(3, 18):  # 3～17列目
                    value = self._parse_number(self.table_vars[changed_row][col].get())
                    period_sum += value
                self.table_vars[changed_row][18].set(self._format_number(period_sum))

                # 2. 同じ工事区分の「計」行の該当年度期間列を更新
                total_row = section_info['total_row']
                start_row = section_info['start']
                end_row = section_info['end']

                col_sum = 0
                for r in range(start_row, end_row + 1):
                    value = self._parse_number(self.table_vars[r][changed_col].get())
                    col_sum += value
                self.table_vars[total_row][changed_col].set(self._format_number(col_sum))

                # 3. 「計」行の年度期間合計を計算（3～17列の合計 → 18列目）
                total_period_sum = 0
                for col in range(3, 18):
                    value = self._parse_number(self.table_vars[total_row][col].get())
                    total_period_sum += value
                self.table_vars[total_row][18].set(self._format_number(total_period_sum))

                # 4. 直接工事計の該当年度期間列を更新
                grand_total_row = WORK_SECTIONS['grand_total']['row']
                grand_col_sum = 0
                for sec_name, sec_info in WORK_SECTIONS.items():
                    if sec_name == 'grand_total' or sec_name == 'operation_cost':
                        continue
                    total_row = sec_info['total_row']
                    value = self._parse_number(self.table_vars[total_row][changed_col].get())
                    grand_col_sum += value
                self.table_vars[grand_total_row][changed_col].set(self._format_number(grand_col_sum))

                # 5. 直接工事計の年度期間合計を計算（3～17列の合計 → 18列目）
                grand_period_sum = 0
                for col in range(3, 18):
                    value = self._parse_number(self.table_vars[grand_total_row][col].get())
                    grand_period_sum += value
                self.table_vars[grand_total_row][18].set(self._format_number(grand_period_sum))

                # print(f"年度期間更新: 行{changed_row}, 列{changed_col}")

            except Exception as e:
                print(f"年度期間変更エラー: {e}")
                import traceback
                traceback.print_exc()

        # ★★★ 【修正】各工事区分と直接工事計の年度期間集計関数 ★★★
        def calculate_section_period_totals(self):
            """
            各工事区分の「計」行の年度期間列（3-17列）と計列（18列）を集計
            さらに直接工事計の年度期間列と計列を集計
            """
            try:
                # print("\n=== calculate_section_period_totals 開始 ===")

                # ★★★ ステップ1: 各工事区分の「計」行について年度期間を集計 ★★★
                for section_name, section_info in WORK_SECTIONS.items():
                    if section_name in ['grand_total', 'operation_cost', 'general_admin_cost']:
                        continue

                    total_row = section_info['total_row']
                    start_row = section_info['start']
                    end_row = section_info['end']

                    # print(f"\n{section_name} (計行: {total_row}, 範囲: {start_row}-{end_row})")

                    # 各年度期間列（3-17列）について集計
                    for col in range(3, 18):
                        col_sum = 0.0
                        for r in range(start_row, end_row + 1):
                            # 既に100万円単位の値が入っているのでそのまま加算
                            value_str = self.table_vars[r][col].get().strip()
                            if value_str and value_str != "":
                                try:
                                    col_sum += float(value_str)
                                except ValueError:
                                    print(f"    行{r}列{col}: 変換エラー '{value_str}'")
                                    continue

                        # そのまま設定（既に100万円単位）
                        if col_sum > 0:
                            self.table_vars[total_row][col].set(f"{col_sum:.2f}")
                        else:
                            self.table_vars[total_row][col].set("")

                        # 年度範囲を取得
                        year_range = self.year_ranges.get(col, (0, 0))
                        if col_sum > 0:
                            print(f"  列{col} ({year_range[0]}-{year_range[1]}年): {col_sum:.2f}百万円")

                    # 計列（18列目）を集計
                    total_sum = 0.0
                    for col in range(3, 18):
                        value_str = self.table_vars[total_row][col].get().strip()
                        if value_str and value_str != "":
                            try:
                                total_sum += float(value_str)
                            except ValueError:
                                continue

                    if total_sum > 0:
                        self.table_vars[total_row][18].set(f"{total_sum:.2f}")
                    else:
                        self.table_vars[total_row][18].set("")
                    print(f"  合計（18列目）: {total_sum:.2f}百万円")

                # ★★★ ステップ2: 直接工事計の年度期間を集計 ★★★
                grand_total_row = WORK_SECTIONS['grand_total']['row']
                print(f"\n直接工事計 (行: {grand_total_row})")
                # print("--- 年度期間集計開始 ---")

                # 各年度期間列（3-17列）について集計
                for col in range(3, 18):
                    grand_col_sum = 0.0

                    # print(f"\n  列{col}の集計:")
                    # 全ての工事区分の「計」行から該当列を集計
                    for sec_name, sec_info in WORK_SECTIONS.items():
                        if sec_name in ['grand_total', 'operation_cost', 'general_admin_cost']:
                            continue

                        total_row = sec_info['total_row']
                        value_str = self.table_vars[total_row][col].get().strip()
                        if value_str and value_str != "":
                            try:
                                value = float(value_str)
                                grand_col_sum += value
                                if value > 0:
                                    print(f"    {sec_name}計(行{total_row}): {value:.2f}百万円")
                            except ValueError:
                                print(f"    {sec_name}計(行{total_row}): 変換エラー '{value_str}'")
                                continue

                    # 直接工事計の該当列に設定
                    if grand_col_sum > 0:
                        self.table_vars[grand_total_row][col].set(f"{grand_col_sum:.2f}")
                    else:
                        self.table_vars[grand_total_row][col].set("")

                    year_range = self.year_ranges.get(col, (0, 0))
                    # print(f"  → 列{col} ({year_range[0]}-{year_range[1]}年) 合計: {grand_col_sum:.2f}百万円")

                # ★★★ ステップ3: 直接工事計の計列（18列目）を集計 ★★★
                # print("\n--- 計列（18列目）集計 ---")
                grand_total_sum = 0.0

                # 3～17列の合計を計算
                for col in range(3, 18):
                    value_str = self.table_vars[grand_total_row][col].get().strip()
                    if value_str and value_str != "":
                        try:
                            value = float(value_str)
                            grand_total_sum += value
                            if value > 0:
                                year_range = self.year_ranges.get(col, (0, 0))
                                print(f"  列{col} ({year_range[0]}-{year_range[1]}年): {value:.2f}百万円")
                        except ValueError:
                            continue

                # 直接工事計の18列目に設定
                if grand_total_sum > 0:
                    self.table_vars[grand_total_row][18].set(f"{grand_total_sum:.2f}")
                else:
                    self.table_vars[grand_total_row][18].set("")
                # print(f"\n直接工事計 合計（18列目）: {grand_total_sum:.2f}百万円")

                # print("\n=== calculate_section_period_totals 完了 ===\n")

            except Exception as e:
                print(f"工事区分年度期間集計エラー: {e}")
                import traceback
                traceback.print_exc()

        # ========== メソッド定義 ==========
        def recalculate_table(self):
            """表全体を再計算

            計算ロジック:
            1. 各工事区分の「計」行の金額 = 「計」行の比率 × 新築工事推定金額
            2. 各工種の金額 = 工種の比率 × その工事区分の「計」行の金額
            3. 直接工事計の金額 = 各工事区分の「計」行の金額の合計
            4. 維持管理・大規模修繕費用の計算（年度期間集計含む）
            """
            try:
                # print("\n=== recalculate_table 開始 ===")

                # 新築工事推定金額を取得
                base_cost_str = self.new_construction_cost.get()
                base_cost = self._parse_number(base_cost_str)
                print(f"新築工事推定金額: {base_cost:,.0f}円")

                if base_cost <= 0:
                    print("新築工事費が設定されていません")
                    return

                # ステップ1: 各工事区分の「計」行の金額を計算
                section_total_amounts = {}
                # print("\n--- 各工事区分の「計」行の金額計算 ---")

                for section_name, section_info in WORK_SECTIONS.items():
                    if section_name in ['grand_total', 'operation_cost', 'general_admin_cost']:
                        continue

                    total_row = section_info['total_row']
                    ratio_str = self.table_vars[total_row][1].get()
                    ratio = self._parse_number(ratio_str)

                    section_amount = base_cost * (ratio / 100)
                    section_total_amounts[section_name] = section_amount

                    self.table_vars[total_row][2].set(self._format_number(section_amount))
                    print(
                        f"{section_name}計 (行{total_row}): {ratio:.2f}% × {base_cost:,.0f} = {section_amount:,.0f}円")

                # ステップ2: 各工種の金額を計算
                # print("\n--- 各工種の金額計算 ---")

                for section_name, section_info in WORK_SECTIONS.items():
                    if section_name in ['grand_total', 'operation_cost', 'general_admin_cost']:
                        continue

                    start_row = section_info['start']
                    end_row = section_info['end']
                    section_total_amount = section_total_amounts[section_name]

                    # print(f"\n{section_name} (工事区分計: {section_total_amount:,.0f}円)")

                    for r in range(start_row, end_row + 1):
                        distribution_ratio_str = self.table_vars[r][1].get()
                        distribution_ratio = self._parse_number(distribution_ratio_str)

                        item_amount = section_total_amount * (distribution_ratio / 100)
                        self.table_vars[r][2].set(self._format_number(item_amount))

                        component_name = self.table_vars[r][0].get()
                        # print(
                        #     f"  行{r} {component_name}: {distribution_ratio:.2f}% × {section_total_amount:,.0f} = {item_amount:,.0f}円")

                # ステップ3: 直接工事計を計算
                print("\n--- 直接工事計の計算 ---")
                grand_total_row = WORK_SECTIONS['grand_total']['row']
                grand_total_amount = sum(section_total_amounts.values())

                self.table_vars[grand_total_row][2].set(self._format_number(grand_total_amount))
                print(f"直接工事計 (行{grand_total_row}): {grand_total_amount:,.0f}円")

                # 直接工事計の比率も更新
                if base_cost > 0:
                    grand_ratio = (grand_total_amount / base_cost) * 100
                    self.table_vars[grand_total_row][1].set(f"{grand_ratio:.2f}%")
                    print(f"直接工事計の比率: {grand_ratio:.2f}%")

                # 純工事費率を更新
                self.update_net_construction_rate()

                # ステップ4: 維持管理・大規模修繕費用を計算（年度期間集計含む）
                self.calculate_maintenance_and_repair_costs()
                self.update_lcc_pie_chart()  # 円グラフを更新
                print("\n=== 表の再計算が完了しました ===\n")

            except Exception as e:
                print(f"再計算エラー: {e}")
                import traceback
                traceback.print_exc()
        # ★★★ 【追加】円グラフ描画関数 ★★★
        def update_lcc_pie_chart(self):
            """LCCコスト構成の円グラフを更新"""
            try:
                # キャンバスをクリア
                self.lcc_pie_chart_canvas.delete("all")

                # データを取得
                # 1. 再調達価格
                reconstruction_cost_str = self.new_construction_cost.get().replace(',', '').strip()
                reconstruction_cost = float(reconstruction_cost_str) if reconstruction_cost_str else 0.0

                # 2. 修繕工事費（直接工事計の18列目）
                grand_total_row = 42  # 直接工事計の行番号
                repair_cost_str = self.table_vars[grand_total_row][18].get().strip()
                repair_cost_million = float(repair_cost_str) if repair_cost_str else 0.0
                repair_cost = repair_cost_million * 1_000_000  # 100万円単位を円に変換

                # 3. 運用費（43行目の18列目）
                operation_row = 43
                operation_cost_str = self.table_vars[operation_row][18].get().strip()
                operation_cost_million = float(operation_cost_str) if operation_cost_str else 0.0
                operation_cost = operation_cost_million * 1_000_000  # 100万円単位を円に変換

                # ★★★ 【追加】4. 一般管理費（44行目の18列目） ★★★
                admin_cost_row = 44
                admin_cost_str = self.table_vars[admin_cost_row][18].get().strip()
                admin_cost_million = float(admin_cost_str) if admin_cost_str else 0.0
                admin_cost = admin_cost_million * 1_000_000  # 100万円単位を円に変換

                # 合計を計算（一般管理費を追加）
                total_cost = reconstruction_cost + repair_cost + operation_cost + admin_cost

                print(f"\n=== 円グラフデータ ===")
                print(f"再調達価格: {reconstruction_cost:,.0f}円")
                print(f"修繕工事費: {repair_cost:,.0f}円")
                print(f"運用費: {operation_cost:,.0f}円")
                print(f"一般管理費: {admin_cost:,.0f}円")
                print(f"合計: {total_cost:,.0f}円")

                if total_cost <= 0:
                    # データがない場合はメッセージを表示
                    self.lcc_pie_chart_canvas.create_text(
                        150, 150,
                        text="データがありません",
                        font=("Arial", 12),
                        fill="gray"
                    )
                    # 凡例をクリア
                    self.lcc_legend_label1.config(text="■ 再調達価格: 0円")
                    self.lcc_legend_label2.config(text="■ 修繕工事費: 0円")
                    self.lcc_legend_label3.config(text="■ 運用費: 0円")
                    self.lcc_legend_label4.config(text="■ 一般管理費: 0円")  # ★★★ 追加
                    return

                # 円グラフの設定
                center_x, center_y = 150, 150
                radius = 100

                # ★★★ 【修正】色の設定（一般管理費用の色を追加）★★★
                colors = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000"]  # 青、オレンジ、グレー、黄色

                # データと色を対応付け（一般管理費を追加）
                data = [
                    ("再調達価格", reconstruction_cost, colors[0]),
                    ("修繕工事費", repair_cost, colors[1]),
                    ("運用費", operation_cost, colors[2]),
                    ("一般管理費", admin_cost, colors[3])  # ★★★ 追加
                ]

                # 円グラフを描画
                start_angle = 0
                for label, value, color in data:
                    if value > 0:
                        # 角度を計算（度数法）
                        extent_angle = (value / total_cost) * 360

                        # 円弧を描画
                        self.lcc_pie_chart_canvas.create_arc(
                            center_x - radius, center_y - radius,
                            center_x + radius, center_y + radius,
                            start=start_angle,
                            extent=extent_angle,
                            fill=color,
                            outline="white",
                            width=2
                        )

                        # パーセンテージを計算
                        percentage = (value / total_cost) * 100

                        # パーセンテージが5%以上の場合のみラベルを表示
                        if percentage >= 5:
                            # ラベルの位置を計算（円弧の中央）
                            label_angle = math.radians(start_angle + extent_angle / 2)
                            label_x = center_x + (radius * 0.7) * math.cos(label_angle)
                            label_y = center_y - (radius * 0.7) * math.sin(label_angle)

                            # ラベルを描画
                            self.lcc_pie_chart_canvas.create_text(
                                label_x, label_y,
                                text=f"{percentage:.1f}%",
                                font=("Arial", 10, "bold"),
                                fill="white"
                            )

                        start_angle += extent_angle

                # 凡例を更新（一般管理費を追加）
                self.lcc_legend_label1.config(
                    text=f"■ 再調達価格: {reconstruction_cost:,.0f}円 ({reconstruction_cost / total_cost * 100:.1f}%)",
                    foreground=colors[0]
                )
                self.lcc_legend_label2.config(
                    text=f"■ 修繕工事費: {repair_cost:,.0f}円 ({repair_cost / total_cost * 100:.1f}%)",
                    foreground=colors[1]
                )
                self.lcc_legend_label3.config(
                    text=f"■ 運用費: {operation_cost:,.0f}円 ({operation_cost / total_cost * 100:.1f}%)",
                    foreground=colors[2]
                )
                # ★★★ 【追加】一般管理費の凡例 ★★★
                self.lcc_legend_label4.config(
                    text=f"■ 一般管理費: {admin_cost:,.0f}円 ({admin_cost / total_cost * 100:.1f}%)",
                    foreground=colors[3]
                )

                print("=== 円グラフ更新完了 ===\n")

            except Exception as e:
                print(f"円グラフ更新エラー: {e}")
                import traceback
                traceback.print_exc()

        def _parse_number(self, value_str):
            """文字列を数値に変換"""
            if not value_str or value_str == "":
                return 0.0
            try:
                return float(value_str.replace(',', '').replace('%', '').strip())
            except ValueError:
                return 0.0

        def _format_number(self, value):
            """数値をカンマ区切りに"""
            if value == 0:
                return ""
            return f"{value:,.0f}"

        def update_cell_value(self, row, col, value):
            """セルに値を設定して再計算"""
            try:
                formatted_value = self._format_number(value) if isinstance(value, (int, float)) else str(value)
                self.table_vars[row][col].set(formatted_value)
                self.recalculate_table()
            except Exception as e:
                print(f"セル更新エラー: {e}")

        def reset_ratios(self):
            """比率を初期値に戻す"""
            try:
                for row, ratio in DEFAULT_RATIOS.items():
                    self.table_vars[row][1].set(f"{ratio:.2f}%")
                self.current_ratios = DEFAULT_RATIOS.copy()
                self.recalculate_table()
                print("比率を初期値に戻しました")
            except Exception as e:
                print(f"比率リセットエラー: {e}")

        def save_ratios(self):
            """現在の比率を保存"""
            try:
                self.current_ratios = {}
                for row in DEFAULT_RATIOS.keys():
                    self.current_ratios[row] = self._parse_number(self.table_vars[row][1].get())
                print(f"比率を保存しました: {len(self.current_ratios)}件")
                import tkinter.messagebox as messagebox
                messagebox.showinfo("保存完了", "現在の比率を保存しました")
            except Exception as e:
                print(f"比率保存エラー: {e}")

        def load_ratios(self):
            """保存した比率を読み込む"""
            try:
                if not self.current_ratios:
                    import tkinter.messagebox as messagebox
                    messagebox.showwarning("警告", "保存された比率がありません")
                    return

                for row, ratio in self.current_ratios.items():
                    self.table_vars[row][1].set(f"{ratio:.2f}%")
                self.recalculate_table()
                import tkinter.messagebox as messagebox
                messagebox.showinfo("読込完了", "保存した比率を読み込みました")
            except Exception as e:
                print(f"比率読込エラー: {e}")

        def update_all_lcc_calculations(self):
            """全体を再計算"""
            try:
                if not hasattr(self, '_ratios_initialized'):
                    for row, ratio in DEFAULT_RATIOS.items():
                        self.table_vars[row][1].set(f"{ratio:.2f}%")
                    self._ratios_initialized = True


                self.recalculate_table()

            except Exception as e:
                print(f"全体計算エラー: {e}")

        def update_net_construction_rate(self):
            """純工事費率と金額を計算"""
            try:
                # print("=== update_net_construction_rate 開始 ===")

                # 直接工事計の比率を取得（42行目、1列目）
                direct_work_ratio_str = self.table_vars[42][1].get()
                direct_work_ratio = self._parse_number(direct_work_ratio_str)
                # print(f"直接工事計の比率: {direct_work_ratio_str} → {direct_work_ratio}")

                # 共通仮設費率を取得
                common_temp_rate_str = self.common_temporary_cost_rate.get()
                common_temp_rate = self._parse_number(common_temp_rate_str)
                # print(f"共通仮設費率: {common_temp_rate_str} → {common_temp_rate}")

                # 純工事費率 = 直接工事計の比率 + 共通仮設費率
                net_construction_rate = direct_work_ratio + common_temp_rate

                # 純工事費率を設定
                self.net_construction_cost_rate.set(f"{net_construction_rate:.2f}")
                # print(f"純工事費率: {net_construction_rate:.2f}%")

                # 新築工事推定金額を取得
                base_cost_str = self.new_construction_cost.get()
                base_cost = self._parse_number(base_cost_str)
                # print(f"新築工事推定金額: {base_cost_str} → {base_cost}")

                # 各項目の金額を計算
                if base_cost > 0:
                    #直接工事費金額
                    direct_amount = base_cost * (direct_work_ratio / 100)
                    formatted_direct = self._format_number(direct_amount)
                    self.direct_construction_cost_amount.set(formatted_direct)
                    self.direct_construction_cost_rate.set(f"{direct_work_ratio:.2f}")
                    print(f"直接工事費: 率={direct_work_ratio:.2f}%, 金額={direct_amount} → {formatted_direct}")

                    # 純工事費金額
                    net_amount = base_cost * (net_construction_rate / 100)
                    formatted_net = self._format_number(net_amount)
                    self.net_construction_cost_amount.set(formatted_net)
                    print(f"純工事費金額: {net_amount} → {formatted_net}")

                    # 現場管理費金額
                    site_management_rate_str = self.site_management_cost_rate.get()
                    site_management_rate = self._parse_number(site_management_rate_str)
                    site_management_amount = base_cost * (site_management_rate / 100)
                    formatted_site = self._format_number(site_management_amount)
                    self.site_management_cost_amount.set(formatted_site)
                    print(f"現場管理費: 率={site_management_rate}%, 金額={site_management_amount} → {formatted_site}")

                    # 共通仮設費金額
                    common_temp_amount = base_cost * (common_temp_rate / 100)
                    formatted_common = self._format_number(common_temp_amount)
                    self.common_temporary_cost_amount.set(formatted_common)
                    print(f"共通仮設費金額: {common_temp_amount} → {formatted_common}")

                    # 一般管理費等金額
                    general_admin_rate_str = self.general_admin_fee_rate.get()
                    general_admin_rate = self._parse_number(general_admin_rate_str)
                    general_admin_amount = base_cost * (general_admin_rate / 100)
                    formatted_admin = self._format_number(general_admin_amount)
                    self.general_admin_fee_amount.set(formatted_admin)
                    print(f"一般管理費等: 率={general_admin_rate}%, 金額={general_admin_amount} → {formatted_admin}")
                else:
                    # 新築工事費がない場合は0
                    self.direct_construction_cost_amount.set("0")
                    self.direct_construction_cost_rate.set("0.00")
                    self.net_construction_cost_amount.set("0")
                    self.site_management_cost_amount.set("0")
                    self.common_temporary_cost_amount.set("0")
                    self.general_admin_fee_amount.set("0")
                    print("新築工事費が0のため、全て0に設定")

                print("=== update_net_construction_rate 完了 ===\n")

            except Exception as e:
                print(f"純工事費率計算エラー: {e}")
                import traceback
                traceback.print_exc()

        # メソッドをクラスに追加
        self.recalculate_table = lambda: recalculate_table(self)
        self._parse_number = lambda value_str: _parse_number(self, value_str)
        self._format_number = lambda value: _format_number(self, value)
        self.update_cell_value = lambda row, col, value: update_cell_value(self, row, col, value)
        self.reset_ratios = lambda: reset_ratios(self)
        self.save_ratios = lambda: save_ratios(self)
        self.load_ratios = lambda: load_ratios(self)
        self.update_all_lcc_calculations = lambda: update_all_lcc_calculations(self)
        self.update_net_construction_rate = lambda: update_net_construction_rate(self)
        self.on_ratio_change = lambda row: on_ratio_change(self, row)
        self.on_period_change = lambda row, col: on_period_change(self, row, col)
        self.update_lcc_year_marker = lambda elapsed_years: update_lcc_year_marker(self, elapsed_years)
        self.calculate_operation_cost = lambda: calculate_operation_cost(self)
        self.calculate_maintenance_and_repair_costs = lambda: calculate_maintenance_and_repair_costs(self)
        self.calculate_section_period_totals = lambda: calculate_section_period_totals(self)
        self.update_lcc_pie_chart = lambda: update_lcc_pie_chart(self)
        self.calculate_total_operation_cost = lambda: calculate_total_operation_cost(self)
        self.calculate_general_admin_cost = lambda operation_cost_monthly=None: calculate_general_admin_cost(self, operation_cost_monthly)


        # ========== 変更監視の設定 ==========
        # 再調達価格の変更を監視
        if hasattr(self, 'new_construction_cost'):
            try:
                for trace_id in self.new_construction_cost.trace_info():
                    self.new_construction_cost.trace_remove(trace_id[0], trace_id[1])
            except:
                pass
            self.new_construction_cost.trace_add('write', lambda *args: (
                self.update_all_lcc_calculations(),
                self.update_net_construction_rate(),
                self.calculate_general_admin_cost()
            ))
            # print("new_construction_costの変更監視を設定しました")

        # 人件費の変更を監視
        if hasattr(self, 'lcc_personnel_cost'):
            def on_personnel_change(*args):
                print("★★★ 人件費が変更されました ★★★")
                self.calculate_total_operation_cost()

            # 既存のtraceを削除（重要）
            try:
                for trace_id in self.lcc_personnel_cost.trace_info():
                    self.lcc_personnel_cost.trace_remove(trace_id[0], trace_id[1])
            except:
                pass

            self.lcc_personnel_cost.trace_add('write', on_personnel_change)
            # print("lcc_personnel_costの変更監視を設定しました")
        else:
            print("警告: lcc_personnel_cost が存在しません")

        # アメニティ費の変更を監視
        if hasattr(self, 'lcc_amenity_cost'):
            try:
                for trace_id in self.lcc_amenity_cost.trace_info():
                    self.lcc_amenity_cost.trace_remove(trace_id[0], trace_id[1])
            except:
                pass
            self.lcc_amenity_cost.trace_add('write', lambda *args: self.calculate_total_operation_cost())  # ← 1つだけ
            # print("lcc_amenity_costの変更監視を設定しました")

        # OTA・広告費の変更を監視
        if hasattr(self, 'lcc_ota_ad_cost'):
            try:
                for trace_id in self.lcc_ota_ad_cost.trace_info():
                    self.lcc_ota_ad_cost.trace_remove(trace_id[0], trace_id[1])
            except:
                pass
            self.lcc_ota_ad_cost.trace_add('write', lambda *args: self.calculate_total_operation_cost())  # ← 1つだけ
            # print("lcc_ota_ad_costの変更監視を設定しました")

        # 食材費の変更を監視
        if hasattr(self, 'lcc_food_cost'):
            try:
                for trace_id in self.lcc_food_cost.trace_info():
                    self.lcc_food_cost.trace_remove(trace_id[0], trace_id[1])
            except:
                pass
            self.lcc_food_cost.trace_add('write', lambda *args: self.calculate_total_operation_cost())  # ← 1つだけ
            # print("lcc_food_costの変更監視を設定しました")

        # リネン費の変更を監視
        if hasattr(self, 'lcc_linen_cost'):
            try:
                for trace_id in self.lcc_linen_cost.trace_info():
                    self.lcc_linen_cost.trace_remove(trace_id[0], trace_id[1])
            except:
                pass
            self.lcc_linen_cost.trace_add('write', lambda *args: self.calculate_total_operation_cost())  # ← 1つだけ
            # print("lcc_linen_costの変更監視を設定しました")

        # 共通仮設費率の変更を監視
        if hasattr(self, 'common_temporary_cost_rate'):
            try:
                for trace_id in self.common_temporary_cost_rate.trace_info():
                    self.common_temporary_cost_rate.trace_remove(trace_id[0], trace_id[1])
            except:
                pass
            self.common_temporary_cost_rate.trace_add('write', lambda *args: (
                self.recalculate_table(),
                self.update_net_construction_rate()
            ))
            # print("common_temporary_cost_rateの変更監視を設定しました")

        # 現場管理費率の変更を監視
        if hasattr(self, 'site_management_cost_rate'):
            try:
                for trace_id in self.site_management_cost_rate.trace_info():
                    self.site_management_cost_rate.trace_remove(trace_id[0], trace_id[1])
            except:
                pass
            self.site_management_cost_rate.trace_add('write', lambda *args: self.update_net_construction_rate())
            # print("site_management_cost_rateの変更監視を設定しました")

        # 一般管理費等率の変更を監視
        if hasattr(self, 'general_admin_fee_rate'):
            try:
                for trace_id in self.general_admin_fee_rate.trace_info():
                    self.general_admin_fee_rate.trace_remove(trace_id[0], trace_id[1])
            except:
                pass
            self.general_admin_fee_rate.trace_add('write', lambda *args: self.update_net_construction_rate())
            # print("general_admin_fee_rateの変更監視を設定しました")

        # 延床面積の変更を監視
        if hasattr(self, 'lcc_building_area'):
            try:
                for trace_id in self.lcc_building_area.trace_info():
                    self.lcc_building_area.trace_remove(trace_id[0], trace_id[1])
            except:
                pass
            self.lcc_building_area.trace_add('write', lambda *args: self.calculate_operation_cost())
            # print("lcc_building_areaの変更監視を設定しました")

        # エネルギー単価の変更を監視
        if hasattr(self, 'lcc_energy_unit_cost'):
            try:
                for trace_id in self.lcc_energy_unit_cost.trace_info():
                    self.lcc_energy_unit_cost.trace_remove(trace_id[0], trace_id[1])
            except:
                pass
            self.lcc_energy_unit_cost.trace_add('write', lambda *args: self.calculate_operation_cost())
            # print("lcc_energy_unit_costの変更監視を設定しました")

        # 水道単価の変更を監視
        if hasattr(self, 'lcc_water_unit_cost'):
            try:
                for trace_id in self.lcc_water_unit_cost.trace_info():
                    self.lcc_water_unit_cost.trace_remove(trace_id[0], trace_id[1])
            except:
                pass
            self.lcc_water_unit_cost.trace_add('write', lambda *args: self.calculate_operation_cost())
            # print("lcc_water_unit_costの変更監視を設定しました")

        # 年間維持管理比率の変更を監視
        if hasattr(self, 'lcc_annual_maintenance_rate'):
            try:
                for trace_id in self.lcc_annual_maintenance_rate.trace_info():
                    self.lcc_annual_maintenance_rate.trace_remove(trace_id[0], trace_id[1])
            except:
                pass
            self.lcc_annual_maintenance_rate.trace_add('write',
                                                       lambda *args: self.calculate_maintenance_and_repair_costs())
            # print("lcc_annual_maintenance_rateの変更監視を設定しました")

        # 大規模修繕費率の変更を監視
        if hasattr(self, 'lcc_major_repair_rate'):
            try:
                for trace_id in self.lcc_major_repair_rate.trace_info():
                    self.lcc_major_repair_rate.trace_remove(trace_id[0], trace_id[1])
            except:
                pass
            self.lcc_major_repair_rate.trace_add('write', lambda *args: self.calculate_maintenance_and_repair_costs())
            # print("lcc_major_repair_rateの変更監視を設定しました")

        # 大規模修繕周期の変更を監視
        if hasattr(self, 'lcc_major_repair_interval'):
            try:
                for trace_id in self.lcc_major_repair_interval.trace_info():
                    self.lcc_major_repair_interval.trace_remove(trace_id[0], trace_id[1])
            except:
                pass
            self.lcc_major_repair_interval.trace_add('write',
                                                     lambda *args: self.calculate_maintenance_and_repair_costs())
            # print("lcc_major_repair_intervalの変更監視を設定しました")

        # 固定資産税率の変更を監視
        if hasattr(self, 'lcc_tax_rate'):
            try:
                for trace_id in self.lcc_tax_rate.trace_info():
                    self.lcc_tax_rate.trace_remove(trace_id[0], trace_id[1])
            except:
                pass
            self.lcc_tax_rate.trace_add('write', lambda *args: self.calculate_general_admin_cost())
            # print("lcc_tax_rateの変更監視を設定しました")

        # 保険料率の変更を監視
        if hasattr(self, 'lcc_insurance_rate'):
            try:
                for trace_id in self.lcc_insurance_rate.trace_info():
                    self.lcc_insurance_rate.trace_remove(trace_id[0], trace_id[1])
            except:
                pass
            self.lcc_insurance_rate.trace_add('write', lambda *args: self.calculate_general_admin_cost())
            # print("lcc_insurance_rateの変更監視を設定しました")

        # ========== 初回計算 ==========
        self.update_all_lcc_calculations()
        # 工事費率の金額も初回計算
        self.update_net_construction_rate()
        # 運用費の初回計算
        self.calculate_operation_cost()
        # 維持管理・大規模修繕費用の初回計算（年度期間集計含む）
        self.calculate_maintenance_and_repair_costs()
        # 円グラフの初回描画
        self.update_lcc_pie_chart()
        # 運営費合計の初回計算
        self.calculate_total_operation_cost()
        # 一般管理費の初回計算
        self.calculate_general_admin_cost()

        # 経過年数マーカーの初期表示
        if hasattr(self, 'date_difference'):
            try:
                days_diff = self.date_difference.get()
                elapsed_years = days_diff / 365.25
                self.update_lcc_year_marker(elapsed_years)
            except:
                self.update_lcc_year_marker(0)
        # ========== LCC修繕計画表システム終了 ==========

        # ========== LCC分析設定 ==========
        ttk.Label(main_frame, text="LCC分析設定", font=sub_title_font).grid(
            row=row, column=0, columnspan=3, pady=(15, 10), sticky=tk.W)
        row += 1

        lcc_settings_frame = ttk.Frame(main_frame)
        lcc_settings_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))

        lcc_settings_data_left = [
            ("分析期間:", self.lcc_analysis_years, "年"),
            ("割引率:", self.lcc_discount_rate, "%"),
            ("年間維持管理費率:", self.lcc_annual_maintenance_rate, "% (再調達価格比)"),
            ("大規模修繕費率:", self.lcc_major_repair_rate, "% (再調達価格比)"),
            ("大規模修繕周期:", self.lcc_major_repair_interval, "年ごと"),
            ("エネルギー単価:", self.lcc_energy_unit_cost, "円/m2/月 電力、ガス、油 "),
            ("水道単価:", self.lcc_water_unit_cost, "円/m2/月"),
            ("保険料率:", self.lcc_insurance_rate, "% (再調達価格の年間)"),
            ("固定資産税率:", self.lcc_tax_rate, "% (再調達価格の年間)"),
        ]

        # 左列のフレーム作成
        left_settings_frame = ttk.Frame(lcc_settings_frame)
        left_settings_frame.grid(row=0, column=0, sticky=(tk.W, tk.N), padx=(0, 30))

        # 左列の項目を配置
        for i, (label, var, unit) in enumerate(lcc_settings_data_left):
            ttk.Label(left_settings_frame, text=label, width=20).grid(row=i, column=0, sticky=tk.W, pady=3)
            ttk.Entry(left_settings_frame, textvariable=var, width=15).grid(
                row=i, column=1, sticky=tk.W, padx=5, pady=3)
            ttk.Label(left_settings_frame, text=unit).grid(row=i, column=2, sticky=tk.W, pady=3)

        # ★★★ 【追加】右列のフレーム作成（運営費セクション） ★★★
        right_settings_frame = ttk.Frame(lcc_settings_frame)
        right_settings_frame.grid(row=0, column=1, sticky=(tk.W, tk.N), padx=(30, 0))

        # 運営費（合計）- 読み取り専用
        ttk.Label(right_settings_frame, text="運営費合計:", width=20, font=("", 10, "bold")).grid(
            row=0, column=0, sticky=tk.W, pady=3)
        ttk.Entry(right_settings_frame, textvariable=self.lcc_operation_cost, width=15,
                  state='readonly', justify=tk.RIGHT).grid(
            row=0, column=1, sticky=tk.W, padx=5, pady=3)
        ttk.Label(right_settings_frame, text="円/月", font=("", 10, "bold")).grid(
            row=0, column=2, sticky=tk.W, pady=3)

        # 運営費内訳データ
        operation_cost_details = [
            ("  人件費:", self.lcc_personnel_cost, "円/月"),
            ("  アメニティ費:", self.lcc_amenity_cost, "円/月"),
            ("  OTA・広告費:", self.lcc_ota_ad_cost, "円/月"),
            ("  食材費:", self.lcc_food_cost, "円/月"),
            ("  リネン費:", self.lcc_linen_cost, "円/月"),
        ]

        # 運営費内訳を配置（1行目から開始）
        for i, (label, var, unit) in enumerate(operation_cost_details):
            ttk.Label(right_settings_frame, text=label, width=20).grid(
                row=i + 1, column=0, sticky=tk.W, pady=3)
            ttk.Entry(right_settings_frame, textvariable=var, width=15, justify=tk.RIGHT).grid(
                row=i + 1, column=1, sticky=tk.W, padx=5, pady=3)
            ttk.Label(right_settings_frame, text=unit).grid(
                row=i + 1, column=2, sticky=tk.W, pady=3)

        row += 1

    def update_construction_unit_price(self, *args):
        """
        選択された建物構造と仕様グレードに基づき、
        self.structure_unit_costs から単価を取得し、lcc_construction_unit_priceを動的に更新する。
        """
        try:
            # 建物構造と仕様グレードの値を取得
            structure = self.structure_type.get()
            spec_grade = self.selected_spec_grade.get()


            # 構造とグレードのインデックスを取得
            structure_idx = self.structure_index_map.get(structure)
            grade_idx = self.grade_index_map.get(spec_grade)

            if structure_idx is not None and grade_idx is not None:
                # 構造の行データを取得 (例: ("RC造", [StringVar, StringVar, ...]))
                _, cost_vars = self.structure_unit_costs[structure_idx]

                # 該当するグレードの tk.StringVar を取得
                price_var = cost_vars[grade_idx]

                # tk.StringVar の値を取得
                base_price_str = price_var.get()


                # lcc_construction_unit_price に設定
                try:
                    # 文字列を数値に変換
                    base_price_int = int(float(base_price_str))

                    # ★★★ 【変更箇所1】建設地域掛け率を取得 ★★★
                    region_rate = self.get_region_rate()  # 0.75〜1.00の値
                    print(f"5. 建設地域係数: {region_rate}")

                    # ★★★ 【追加】地域掛け率表示用変数を更新 ★★★
                    self.lcc_region_rate_display.set(f"建設地域係数: {region_rate:.2f}")
                    # ★★★ 【変更箇所2】建築基準単価に地域掛け率を適用して表示 ★★★
                    adjusted_unit_price = base_price_int * region_rate
                    # カンマ区切りで表示
                    self.lcc_construction_unit_price.set(f"{int(adjusted_unit_price):,}")

                    print(f"6. 調整後単価: {int(adjusted_unit_price):,}")

                    # 建物延床面積を取得し、数値に変換
                    area_str = self.lcc_building_area.get()
                    area_clean = area_str.replace(',', '').strip()  # カンマを除去

                    if area_clean:  # 面積が入力されている場合のみ計算
                        area = float(area_clean)

                        # ★★★ 【変更箇所3】再調達価格計算 ★★★
                        # 再調達価格 = 建築基準単価(既に地域掛け率適用済み) × 建物延床面積
                        estimated_cost = adjusted_unit_price * area

                        # 新築工事推定金額をカンマ区切りでtk.StringVarに設定 (整数表示)
                        self.new_construction_cost.set(f"{estimated_cost:,.0f}")
                        print(f"7. 再調達価格: {estimated_cost:,.0f}")
                    else:
                        self.new_construction_cost.set("0")

                except ValueError as ve:
                    self.lcc_construction_unit_price.set("0")
                    self.new_construction_cost.set("0")  # エラー時も初期化
                    print(f"警告: 取得した単価 '{base_price_str}' または面積が数値ではありません: {ve}")
            else:
                # 構造またはグレードの選択が無効な場合
                print(f"警告: structure_idx={structure_idx}, grade_idx={grade_idx}")
                self.lcc_construction_unit_price.set("0")
                self.new_construction_cost.set("0")

        except Exception as e:
            self.lcc_construction_unit_price.set("0")
            self.new_construction_cost.set("0")
            print(f"建築基準単価の更新中にエラーが発生しました: {e}")
            import traceback
            traceback.print_exc()

    def calculate_total_operation_cost(self):
        """
        運営費の各内訳を合計して、運営費（合計）に反映
        """
        try:
            # 各内訳を取得
            personnel = self._parse_number(self.lcc_personnel_cost.get())
            amenity = self._parse_number(self.lcc_amenity_cost.get())
            ota_ad = self._parse_number(self.lcc_ota_ad_cost.get())
            food = self._parse_number(self.lcc_food_cost.get())
            linen = self._parse_number(self.lcc_linen_cost.get())

            # 合計を計算
            total = personnel + amenity + ota_ad + food + linen

            # 運営費（合計）に設定（カンマ区切り）
            self.lcc_operation_cost.set(f"{total:,.0f}")

            print(f"運営費合計更新: {total:,.0f}円/月")
            self.calculate_general_admin_cost(operation_cost_monthly=total)

        except Exception as e:
            print(f"運営費合計計算エラー: {e}")
            import traceback
            traceback.print_exc()
    # LCC設定のタブに関する項目終了

if __name__ == "__main__":
    root = tk.Tk()
    app = BuildingCostCalculator(root)
    root.mainloop()
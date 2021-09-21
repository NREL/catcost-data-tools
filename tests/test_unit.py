import os
import pandas as pd
import unittest

# load project module
from context import ccdt


DEBUG = False

TEST_INDEX = 0
TEST_XLSX = [f for f in os.listdir(os.path.join(os.getcwd(), 'tests', 'data')) if 'xlsx' in f if f[0] != '~']

if DEBUG: print(TEST_XLSX[TEST_INDEX])


class TestGetMatLib(unittest.TestCase):
    def setUp(self, debug=False):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        self.result = ccdt.get_materials_lib(xlsx)
        if debug: print(self.result.head())

    def test_get_materials_lib_df(self):
        """
        Test get_materials_lib returns a DataFrame
        """
        self.assertIsInstance(self.result, pd.DataFrame)
        
    def test_get_materials_lib_some_data(self):
        """
        Test get_materials_lib returns some data
        """
        self.assertGreater(len(self.result), 0)
        
    def test_get_materials_lib_good_data(self):
        """
        Test get_materials_lib returns good data
        """
        self.assertIn('Material Name', self.result.columns)


# class TestGetICIS  # ignore
# class TestRenameTons  # ignore


class TestGenId(unittest.TestCase):
    def test_uuid(self):
        """
        Test gen_id returns a uuid
        """
        result = ccdt.gen_id()
        self.assertEqual(len(result), 36)


class TestDateToStr(unittest.TestCase):
    def setUp(self, debug=False):
        from pandas import Timestamp
        ts = Timestamp(year=2021, month=2, day=2, hour=12, minute=4, second=7)
        data = {'Quote Access Date': ts}
        self.result = ccdt.date_to_str(data)
        if debug: print(self.result)

    def test_date_to_str_len(self):
        """
        Test date_to_str returns two elements
        """
        self.assertEqual(len(self.result), 2)
    
    def test_date_to_str_date(self):
        """
        Test date_to_str returns a valid date string
        """
        from re import match
        result = match(r'\d{2}/\d{2}/\d{4}', self.result[1])
        self.assertEqual(result.span(), (0,10))


class TestMakePriceDict(unittest.TestCase):
    def setUp(self, debug=False):
        """
        Makes entry to feed to make_price_dict (via mat_lib_dict)
        """
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        mat_lib_df = pd.read_excel(xlsx, sheet_name="Materials Library",
                                   skiprows=14, skipfooter=0)
        mat_lib_df = mat_lib_df.rename(columns={"Material Name":"name","Material Type":"type",
                                                "MW (g/mol)":"molecularWeight",
                                                "Density (g/mL)":"density",
                                                "Concentration (%)":"concentration",
                                                "Lab Units":"lab_scale_units",
                                                "Bulk Quote Units":"bulk_quote_units"})
        mat_lib_df = mat_lib_df.drop(["Notes","Basis Cell"], 1)

        mat_lib_dict = mat_lib_df.to_dict('records')
        entry = mat_lib_dict[0]
        if debug: print(entry)
        self.result = ccdt.make_price_dict(entry)
        if debug: print(self.result.keys())
    
    def test_make_price_dict_type(self):
        self.assertIsInstance(self.result, dict)

    def test_make_price_dict_len(self):
        self.assertGreater(len(self.result), 0)

    def test_make_price_dict_key(self):
        self.assertIn('year', self.result.keys())


class TestMaterialsToJson(unittest.TestCase):
    # Better in integration testing
    def setUp(self, debug=False):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        self.json_path = os.path.join('tests', 'data', 'test_mat_to_json.json')
        self.result = ccdt.materials_to_json(xlsx, self.json_path)
        if debug: print(type(self.result))

    def test_materials_to_json_file_exists(self):
        self.assertTrue(os.path.isfile(self.json_path))

    def test_materials_to_json_file_has_data(self):
        self.assertGreater(os.stat(self.json_path).st_size, 0)

    def test_materials_to_json_str(self):
        self.assertIsInstance(self.result, str)

    def test_materials_to_json_dict_data(self):
        self.assertGreater(len(self.result), 0)

    def tearDown(self):
        if os.path.exists(self.json_path):
            os.remove(self.json_path)


class TestGetEquipment(unittest.TestCase):
    def setUp(self):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        self.result = ccdt.get_equipment(xlsx)
    
    def test_get_equipment_df(self):
        self.assertIsInstance(self.result, pd.DataFrame)

    def test_get_equipment_data(self):
        self.assertGreater(len(self.result), 0)


class TestRowManipulations(unittest.TestCase):
    def setUp(self, debug=False):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        df = ccdt.get_equipment(xlsx)
        self.row = df.iloc[10]
        if debug: print(self.row)

    def test_split_nc(self):
        result = ccdt.split_nc(self.row)
        self.assertIsInstance(result, float)

    def test_rename_func_type(self):
        result = ccdt.rename_func_type(self.row)
        self.assertIsInstance(result, str)


# TODO: split into separate test_integration.py file
class TestEquipToJson(unittest.TestCase):
    # Better in integration testing
    def setUp(self, debug=False):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        self.json_path = os.path.join('tests', 'data', 'test_equip_to_json.json')
        self.result = ccdt.materials_to_json(xlsx, self.json_path)
        if debug: print(type(self.result))

    def test_equip_to_json_file_exists(self):
        self.assertTrue(os.path.isfile(self.json_path))

    def test_equip_to_json_file_has_data(self):
        self.assertGreater(os.stat(self.json_path).st_size, 0)

    def test_equip_to_json_str(self):
        self.assertIsInstance(self.result, str)

    def test_equip_to_json_dict_data(self):
        self.assertGreater(len(self.result), 0)

    def tearDown(self):
        if os.path.exists(self.json_path):
            os.remove(self.json_path)


class TestMakePricingBasisLst(unittest.TestCase):
    def setUp(self):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        equip_df = pd.read_excel(xlsx, sheet_name="Equip. Library", skipfooter=1)
        equip_df = equip_df.rename(columns={"Category (not in use)": "category",
                                            "Name": "name", "Year": "year",
                                            "Units for Size, S": "size_unit",
                                            "S lower": "size_min", "S upper": "size_max",
                                            "BM Factor (not in use)": "bm_factor",
                                            "Installation Factor (Garrett)": "installation_factor",
                                            "Note": "note", "Source": "source",
                                            "CEPCI": "cepci","NF Refinery": "nf_refinery",
                                            "Labor Factor": "labor_factor"})
        equip_df = equip_df[equip_df.size_unit.notnull()]
        equip_dict = equip_df.to_dict('records')
        entry = equip_dict[0]
        self.result = ccdt.make_pricing_basis_lst(entry)

    def test_make_pricing_basis_lst(self):
        self.assertIsInstance(self.result, list)

    def test_make_pricing_basis_len(self):
        self.assertGreater(len(self.result), 0)

        
class TestMakeSpentCatTables(unittest.TestCase):
    def setUp(self):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        self.result = ccdt.make_spent_cat_tables(xlsx)

    def test_make_spent_cat_tables_len(self):
        self.assertEqual(len(self.result), 5)

    def test_make_spent_cat_tables_list(self):
        self.assertIsInstance(self.result, list)
        

class TestMakeSpentCatDicts(unittest.TestCase):
    def setUp(self, debug=False):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        tables = ccdt.make_spent_cat_tables(xlsx)
        self.result_support = ccdt.make_support_dict(tables)
        self.result_metal = ccdt.make_metal_dict(tables)
        self.result_hazard = ccdt.make_hazard_dict(tables)
        self.result_density = ccdt.make_density_dict(tables)
    
    def test_make_support_dict_tuple(self):
        self.assertIsInstance(self.result_support, tuple)

    def test_make_support_dict_tuple_len(self):
        self.assertEqual(len(self.result_support), 2)
    
    def test_make_support_dict_0_type(self):
        self.assertIsInstance(self.result_support[0], list)
    
    def test_make_support_dict_1_type(self):
        self.assertIsInstance(self.result_support[1], dict)
        
    def test_make_support_dict_0_data(self):
        self.assertGreater(len(self.result_support[0]), 0)
    
    def test_make_support_dict_1_key(self):
        self.assertIn('loss_of_metal_fixed', self.result_support[1])

    def test_make_metal_dict_tuple(self):
        self.assertIsInstance(self.result_metal, tuple)

    def test_make_metal_dict_tuple_len(self):
        self.assertEqual(len(self.result_metal), 2)
    
    def test_make_metal_dict_0_type(self):
        self.assertIsInstance(self.result_metal[0], list)
    
    def test_make_metal_dict_1_type(self):
        self.assertIsInstance(self.result_metal[1], dict)
        
    def test_make_metal_dict_0_data(self):
        self.assertGreater(len(self.result_metal[0]), 0)
    
    def test_make_metal_dict_1_data(self):
        self.assertGreater(len(self.result_metal[1]), 0)

    def test_make_hazard_dict_type(self):
        self.assertIsInstance(self.result_hazard, list)
    
    def test_make_hazard_dict_data(self):
        self.assertGreater(len(self.result_hazard), 0)

    def test_make_density_dict_type(self):
        self.assertIsInstance(self.result_density, list)
 
    def test_make_density_dict_data(self):
        self.assertGreater(len(self.result_density), 0)
        

class TestMakeDicts(unittest.TestCase):
    def setUp(self):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        tables = ccdt.make_spent_cat_tables(xlsx)
        support_df = tables[0]
        support_df = support_df.rename(columns={'Support':'name', 
                                                'Loss of catalyst solids in use, fixed, %':'loss_of_catalyst_solids_fixed',
                                                'Loss of metal in use, fixed, %':'loss_of_metal_fixed',
                                                'Loss of catalyst solids in use, slurry/fluidized, %':'loss_of_catalyst_solids_slurry',
                                                'Loss of metal in use, slurry/fluidized, %':'loss_of_metal_slurry'})
        support_dict = support_df.to_dict('records')
        entry = support_dict[0]
        self.result_incoming = ccdt.make_incoming_dict(entry)
        self.result_thermal_ox = ccdt.make_thermal_ox_dict(entry)
        self.result_metal_contaminant = ccdt.make_metal_contaminant_dict(entry)
    
    def test_make_incoming_dict_tuple(self):
        self.assertIsInstance(self.result_incoming, tuple)

    def test_make_incoming_dict_tuple_len(self):
        self.assertEqual(len(self.result_incoming), 2)
    
    def test_make_incoming_dict_0_type(self):
        self.assertIsInstance(self.result_incoming[0], dict)
    
    def test_make_incoming_dict_1_type(self):
        self.assertIsInstance(self.result_incoming[1], bool)
        
    def test_make_incoming_dict_0_data(self):
        self.assertGreater(len(self.result_incoming[0]), 0)
    
    def test_make_thermal_ox_dict(self):
        self.assertIsInstance(self.result_thermal_ox, dict)
    
    def test_make_thermal_ox_dict_key(self):
        self.assertIn('baseline', self.result_thermal_ox)
        
    def test_make_thermal_ox_dict_data(self):
        self.assertGreater(len(self.result_thermal_ox), 0)

    def test_make_metal_contaminant_dict(self):
        self.assertIsInstance(self.result_metal_contaminant, dict)
    
    def test_make_metal_contaminant_key(self):
        self.assertIn('baseline', self.result_metal_contaminant)
        
    def test_make_metal_contaminant_data(self):
        self.assertGreater(len(self.result_metal_contaminant), 0)


class TestMakeMetalLossDict(unittest.TestCase):
    def setUp(self):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        tables = ccdt.make_spent_cat_tables(xlsx)
        metal_df = tables[1]
        metal_df = metal_df.rename(columns={'Metal':'name', 
                                            'Refining charge, $/troy oz recovered':'refining_charge',
                                            'Note':'note', 'PGM/Noble (Refining charge yes/no)':'has_refining_charge',
                                            'Spot Price ($)':'spot_price','Unit':'unit',
                                            'Year':'year','Source':'source'})
        metal_dict = metal_df.to_dict('records')
        entry = metal_dict[0]
        self.result = ccdt.make_metal_loss_dict(entry)

    def test_make_metal_loss_dict_tuple(self):
        self.assertIsInstance(self.result, tuple)

    def test_make_metal_loss_dict_tuple_len(self):
        self.assertEqual(len(self.result), 2)
    
    def test_make_metal_loss_dict_0_type(self):
        self.assertIsInstance(self.result[0], dict)
    
    def test_make_metal_loss_dict_1_type(self):
        self.assertIsInstance(self.result[1], bool)
        
    def test_make_metal_loss_dict_0_data(self):
        self.assertGreater(len(self.result[0]), 0)
    
    def test_make_metal_loss_dict_0_key(self):
        self.assertIn('baseline', self.result[0])


class TestMakeHazardDicts(unittest.TestCase):
    def setUp(self):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        tables = ccdt.make_spent_cat_tables(xlsx)
        hazard_df = tables[2]
        hazard_df = hazard_df.rename(columns={'Catalyst Hazard Class':'name','Note':'note'})
        hazard_dict = hazard_df.to_dict('records')
        entry = hazard_dict[0]
        self.result_landfill = ccdt.make_landfill_dict(entry)
        self.result_sale = ccdt.make_sale_dict(entry)
    
    def test_make_landfill_dict_type(self):
        self.assertIsInstance(self.result_landfill, dict)

    def test_make_landfill_dict_key(self):
        self.assertIn('baseline', self.result_landfill.keys())

    def test_make_sale_dict_type(self):
        self.assertIsInstance(self.result_sale, dict)

    def test_make_sale_dict_key(self):
        self.assertIn('baseline', self.result_sale.keys())


class TestSpentCatToJson(unittest.TestCase):
    # Better in integration testing
    def setUp(self, debug=False):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        self.json_path = os.path.join('tests', 'data', 'test_spent_cat_to_json.json')
        self.result = ccdt.spent_cat_to_json(xlsx, self.json_path)
        if debug: print(type(self.result))

    def test_spent_cat_to_json_file_exists(self):
        self.assertTrue(os.path.isfile(self.json_path))

    def test_spent_cat_to_json_file_has_data(self):
        self.assertGreater(os.stat(self.json_path).st_size, 0)

    def test_spent_cat_to_json_str(self):
        self.assertIsInstance(self.result[0], str)

    def test_spent_cat_to_json__data(self):
        self.assertGreater(len(self.result[0]), 0)

    def test_spent_cat_to_json_support_data(self):
        self.assertGreater(len(self.result[1]), 0)

    def test_spent_cat_to_json_metal_data(self):
        self.assertGreater(len(self.result[2]), 0)

    def test_spent_cat_to_json_tuple(self):
        self.assertIsInstance(self.result, tuple)

    def test_spent_cat_to_json_len(self):
        self.assertEqual(len(self.result), 3)

    def tearDown(self):
        if os.path.exists(self.json_path):
            os.remove(self.json_path)


class TestEstimateToJson(unittest.TestCase):
    def setUp(self):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        self.json_path = os.path.join('tests', 'data', 'test_est_to_json.json')
        self.result = ccdt.estimate_to_json(xlsx, self.json_path)

    def test_est_to_json_file_exists(self):
        self.assertTrue(os.path.isfile(self.json_path))

    def test_est_to_json_file_has_data(self):
        self.assertGreater(os.stat(self.json_path).st_size, 0)

    def test_est_to_json_str(self):
        self.assertIsInstance(self.result[0], str)

    def test_est_to_json__data(self):
        self.assertGreater(len(self.result[0]), 0)

    def test_est_to_json_1_data(self):
        self.assertGreater(len(self.result[1]), 0)

    def test_est_to_json_2_data(self):
        self.assertGreater(len(self.result[2]), 0)

    def test_est_to_json_3_data(self):
        self.assertGreaterEqual(len(self.result[3]), 0)

    def test_est_to_json_tuple(self):
        self.assertIsInstance(self.result, tuple)

    def test_est_to_json_len(self):
        self.assertEqual(len(self.result), 4)

    def tearDown(self):
        if os.path.exists(self.json_path):
            os.remove(self.json_path)
    

class TestLocateData(unittest.TestCase):
    def setUp(self, debug=False):
        from xlrd import open_workbook

        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        with open_workbook(xlsx) as wb:
            sheets = wb.sheets()
            for sheet in sheets:
                if sheet.name == '4 Spent Catalyst':
                    spent_cat = sheet
                    break
        row_value = spent_cat.row_values(7)
        if debug: print(row_value)
        self.result_nosens = ccdt.locate_data(row_value, row_value[2], False)
        self.result_sens = ccdt.locate_data(row_value, row_value[2], True)
        if debug: print(self.result_nosens)

    def test_locate_data_sens_tuple(self):
        self.assertIsInstance(self.result_sens, tuple)

    def test_locate_data_sens_len(self):
        self.assertEqual(len(self.result_sens), 2)
        
    def test_locate_data_sens_dict(self):
        self.assertIsInstance(self.result_sens[0], dict)
        
    def test_locate_data_sens_bool(self):
        self.assertIsInstance(self.result_sens[1], bool)
        
    def test_locate_data_nosens_float(self):
        if self.result_nosens != '':  # skip if no entry
            self.assertIsInstance(self.result_nosens, float)


# TODO: suppress running setUp for every test method?
class TestMakeEstEquipLst(unittest.TestCase):
    def setUp(self, debug=False):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        est_id = 'test_id'
        version = '0.0a'
        self.result = ccdt.make_est_equip_lst(xlsx, est_id, version)
        if debug: print(self.result[1:])

    def test_make_est_equip_lst_tuple(self):
        self.assertIsInstance(self.result, tuple)

    def test_make_est_equip_lst_tuple_len(self):
        self.assertEqual(len(self.result), 5)

    def test_make_est_equip_lst_0_data(self):
        self.assertGreater(len(self.result[0]), 0)

    def test_make_est_equip_lst_0_str(self):
        self.assertIsInstance(self.result[0], list)

    def test_make_est_equip_lst_1_str(self):
        self.assertIsInstance(self.result[1], str)

    def test_make_est_equip_lst_2_str(self):
        self.assertIsInstance(self.result[2], str)

    def test_make_est_equip_lst_3_str(self):
        self.assertIsInstance(self.result[3], str)

    def test_make_est_equip_lst_4_flt(self):
        self.assertIsInstance(self.result[4], float)


class TestMakeEstMatLst(unittest.TestCase):
    def setUp(self, debug=False):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        est_id = 'test_id'
        version = '0.0a'
        self.result = ccdt.make_est_mat_lst(xlsx, est_id, version)
        if debug:
            print(self.result)

    def test_make_est_mat_lst_tuple(self):
        self.assertIsInstance(self.result, tuple)

    def test_make_est_mat_lst_tuple_len(self):
        self.assertEqual(len(self.result), 2)

    def test_make_est_mat_lst_0_data(self):
        self.assertGreaterEqual(len(self.result[0]), 0)

    def test_make_est_mat_lst_0_lst(self):
        self.assertIsInstance(self.result[0], list)

    def test_make_est_mat_lst_1_lst(self):
        self.assertIsInstance(self.result[1], list)

    def test_make_est_mat_lst_1_data(self):
        self.assertGreaterEqual(len(self.result[1]), 0)


class TestMakeEstSpentCat(unittest.TestCase):
    def setUp(self, debug=False):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        est_id = 'test_id'
        version = '0.0a'
        self.result = ccdt.make_est_spent_cat(xlsx, est_id, version)
        if debug: print(self.result[1:])

    def test_make_est_spent_cat_tuple(self):
        self.assertIsInstance(self.result, tuple)

    def test_make_est_spent_cat_tuple_len(self):
        self.assertEqual(len(self.result), 5)

    def test_make_est_spent_cat_0_data(self):
        self.assertGreater(len(self.result[0]), 0)

    def test_make_est_spent_cat_0_dict(self):
        self.assertIsInstance(self.result[0], dict)

    def test_make_est_spent_cat_1_dict(self):
        self.assertIsInstance(self.result[1], dict)

    def test_make_est_spent_cat_2_dict(self):
        self.assertIsInstance(self.result[2], dict)

    def test_make_est_spent_cat_3_dict(self):
        self.assertIsInstance(self.result[3], dict)

    def test_make_est_spent_cat_4_dict(self):
        self.assertIsInstance(self.result[4], dict)


class TestMakeEstProcessUtilities(unittest.TestCase):
    def setUp(self, debug=False):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        est_id = 'test_id'
        version = '0.0a'
        basis_unit = pd.read_excel(xlsx, sheet_name='1 Inputs', header=None, usecols='D', skiprows=13, nrows=1).iloc[0,0]
        self.result = ccdt.make_est_process_utilities(xlsx, est_id, version,
                                                      basis_unit)
        if debug: print(self.result[1:])

    def test_make_est_process_utilities_list(self):
        self.assertIsInstance(self.result, list)

    def test_make_est_process_utilities_data(self):
        self.assertGreater(len(self.result), 0)

    def test_make_est_process_utilities_0_data(self):
        self.assertGreater(len(self.result[0]), 0)

    def test_make_est_process_utilities_0_dict(self):
        self.assertIsInstance(self.result[0], dict)


class TestMakeEstCapEx(unittest.TestCase):
    def setUp(self, debug=False):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        est_id = 'test_id'
        version = '0.0a'
        self.result = ccdt.make_est_cap_ex(xlsx, est_id, version)
        if debug: print('len', len(self.result))

    def test_make_est_cap_ex_list(self):
        self.assertIsInstance(self.result, list)

    def test_make_est_cap_ex_data(self):
        self.assertGreater(len(self.result), 0)

    def test_make_est_cap_ex_0_data(self):
        self.assertGreater(len(self.result[0]), 0)

    def test_make_est_cap_ex_0_dict(self):
        self.assertIsInstance(self.result[0], dict)

    def test_make_est_cap_ex_0_key(self):
        self.assertIn('percent_purchase_cost', self.result[0].keys())


class TestMakeEstOpEx(unittest.TestCase):
    def setUp(self, debug=False):
        xlsx = os.path.join('tests', 'data', TEST_XLSX[TEST_INDEX])
        est_id = 'test_id'
        version = '0.0a'
        self.result = ccdt.make_est_op_ex(xlsx, est_id, version)
        if debug: print(self.result)

    def test_make_est_op_ex_list(self):
        self.assertIsInstance(self.result, list)

    def test_make_est_op_ex_data(self):
        self.assertGreater(len(self.result), 0)

    def test_make_est_op_ex_0_data(self):
        self.assertGreater(len(self.result[0]), 0)

    def test_make_est_op_ex_0_dict(self):
        self.assertIsInstance(self.result[0], dict)

    def test_make_est_op_ex_0_key(self):
        self.assertIn('factor', self.result[0].keys())


class TestGetIds(unittest.TestCase):
    def setUp(self, debug=False):
        lib = 'equip_id_dict'
        self.result = ccdt.get_ids(lib)
        if debug: print(self.result)
    
    def test_get_ids_dict(self):
        self.assertIsInstance(self.result, dict)

    def test_get_ids_data(self):
        self.assertGreater(len(self.result), 0)


class TestAddId(unittest.TestCase):

    def setUp(self):
        from shutil import copy2
        # backup all_ids.json for testing
        # will check ~/.catcost-data-tools then catcost-data-tools/catcost_data_tools/default for all_ids.json
        try:
            self.path = os.path.join(os.path.expanduser('~'), '.catcost-data-tools')
            copy2(os.path.join(self.path, 'all_ids.json'), os.path.join(self.path, 'all_ids.json.bak'))
        except FileNotFoundError:
            self.path = os.path.join(os.path.dirname(os.path.dirname(__file__)),
                                    'catcost_data_tools', 'default')
            copy2(os.path.join(self.path, 'all_ids.json'), os.path.join(self.path, 'all_ids.json.bak'))
    

    def test_add_id_file(self):
        from json import loads
        
        lib = 'hazard_id_dict'
        name = 'slinky'
        self.result = ccdt.add_id(lib, name)
        path = os.path.join(self.path, 'all_ids.json')
        with open(path) as f:
            id_dict = loads(f.read())[lib]
        self.assertIn(name, id_dict.keys())


    def tearDown(self):
        from shutil import copy2, move
        # restore all_ids.json
        copy2(os.path.join(self.path, 'all_ids.json.bak'), os.path.join(self.path, 'all_ids.json'))
        os.remove(os.path.join(self.path, 'all_ids.json.bak'))

        
if __name__ == '__main__':
    # open log file
    with open(os.path.join(os.getcwd(), 'unittest.log'), 'a') as f:
        runner = unittest.TextTestRunner(f)
        loader = unittest.TestLoader()
        # iterate over all files in tests/data/ folder
        for TEST_INDEX in range(len(TEST_XLSX)):
            print(TEST_INDEX, TEST_XLSX[TEST_INDEX])
            # get all tests from this file
            tests = loader.discover(os.path.dirname(__file__))
            runner.run(tests)

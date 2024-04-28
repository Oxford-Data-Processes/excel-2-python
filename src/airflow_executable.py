# %%
import os

from series_extraction.excel_loader import ExcelLoader
from series_extraction.excel_cleaner import ExcelCleaner
from series_extraction.table_extractor import TableFinder
from series_extraction.series_extractor import SeriesExtractor
from series_extraction.excel_compatibility_checker import ExcelCompatibilityChecker
from series_extraction.series_iterator import SeriesIterator
from series_extraction.series_mapper import SeriesMapper
from series_extraction.excel_validator import ExcelValidator

from ast_transformation.series_formula_generator import SeriesFormulaGenerator
from ast_transformation.formula_generator import FormulaGenerator
from ast_transformation.formula_evaluator import FormulaEvaluator

from ast_building.formula_parser import FormulaParser
from ast_building.series_implementer import SeriesImplementer

from pipeline_building.series_dependencies_builder import SeriesDependenciesBuilder
from pipeline_building.dag_sorter import DAGSorter

from excel_builder import ExcelBuilder

from excel_checker import ExcelChecker


# %%
data_directory = "/Users/chrislittle/GitHub/speedsheet/excel-2-python/data"

project_name = "vehicle_data"

excel_raw_file_path = os.path.join(
    data_directory, "excel_files_raw", f"{project_name}_raw.xlsx"
)
excel_reduced_filepath = os.path.join(
    data_directory, "excel_files_reduced", f"{project_name}_reduced.xlsx"
)
excel_reduced_clean_filepath = os.path.join(
    data_directory, "excel_files_reduced_clean", f"{project_name}_reduced_clean.xlsx"
)
excel_reduced_clean_series_filepath = os.path.join(
    data_directory,
    "excel_files_reduced_clean_series",
    f"{project_name}_reduced_clean_series.xlsx",
)
excel_reduced_clean_series_python_filepath = os.path.join(
    data_directory,
    "excel_files_reduced_clean_series_python",
    f"{project_name}_reduced_clean_series_python.xlsx",
)

# %%
excel_raw = ExcelLoader.load_file(excel_raw_file_path)
excel_reduced = ExcelLoader.load_file(excel_reduced_filepath)
is_valid = ExcelValidator.validate_excel(excel_reduced)
if not is_valid:
    raise Exception("Excel file is not valid")

excel_reduced_clean = ExcelCleaner.clean_excel(excel_reduced)
ExcelBuilder.create_excel_from_workbook(
    excel_reduced_clean.workbook_with_formulas, excel_reduced_clean_filepath
)

extracted_tables, data = TableFinder.find_tables(excel_reduced_clean)
is_compatible = ExcelCompatibilityChecker.check_file(
    excel_raw, excel_reduced, extracted_tables
)
if not is_compatible:
    print(extracted_tables)
    raise Exception("Excel file is not compatible")

series_dict = SeriesExtractor.extract_series(
    extracted_tables=extracted_tables, data=data
)
series_mapping = SeriesMapper.map_series(series_dict)
series_iterator = SeriesIterator.iterate_series(series_dict)

series_list = [series for series in series_iterator]

series_list_with_formulas = [
    series for series in series_list if series.formulas != [None, None]
]
series_list_with_values = [
    series for series in series_list if series.formulas == [None, None]
]

series_list_new = []
formula_1_ast_series_list = []
ast_generator_dict = {}

for series in series_list_with_formulas:
    formula_1, formula_2 = SeriesFormulaGenerator.adjust_formulas(series.formulas)
    if formula_1 is not None and formula_2 is not None:

        series_implementer = SeriesImplementer(
            series_mapping, sheet_name=series.worksheet.sheet_name
        )

        formula_1_ast = FormulaParser.parse_formula(formula_1)
        formula_1_ast_series = series_implementer.update_ast(formula_1_ast)
        formula_1_ast_series_list.append((series.series_id, formula_1_ast_series))

        formula_2_ast = FormulaParser.parse_formula(formula_2)
        formula_2_ast_series = series_implementer.update_ast(formula_2_ast)

        SeriesFormulaGenerator.process_series_formulas(
            series,
            formula_1_ast_series,
            formula_2_ast_series,
            series_mapping,
            series_dict,
            series_list_new,
        )

        sheet_name = series.worksheet.sheet_name
        series_list_within_sheet = series_dict.get(sheet_name)
        ast_generator = FormulaGenerator.get_ast_generator(
            formula_1_ast_series, formula_2_ast_series, series_list_within_sheet
        )
        ast_generator_dict[series.series_id] = ast_generator


series_list_updated = series_list_new + series_list_with_values

ExcelBuilder.create_excel_from_series(
    series_list_updated, excel_reduced_clean_series_filepath
)
ExcelChecker.excels_are_equivalent(
    excel_reduced_clean_filepath, excel_reduced_clean_series_filepath
)

# %%
series_dependencies = SeriesDependenciesBuilder.build_dependencies(
    formula_1_ast_series_list
)
sorted_dag = DAGSorter.sort_dag(series_dependencies)


# %%
def build_excel_with_python_formulas(
    series_list_with_values,
    ast_generator_dict,
    evaluator,
    excel_reduced_clean_series_python_filepath,
):

    series_list_new_python = []
    for series_id in ast_generator_dict.keys():
        ast_generator = ast_generator_dict[series_id]
        series = evaluator.get_series_from_id(series_id)
        values = evaluator.calculate_series_values(ast_generator, 1, 2)
        series.values = values
        series_list_new_python.append(series)

    series_list_updated_python = series_list_new_python + series_list_with_values
    ExcelBuilder.create_excel_from_series(
        series_list_updated_python,
        excel_reduced_clean_series_python_filepath,
        values_only=True,
    )


# %%
evaluator = FormulaEvaluator(formula_ast=None, series_dict=series_dict)


# %%
build_excel_with_python_formulas(
    series_list_with_values,
    ast_generator_dict,
    evaluator,
    excel_reduced_clean_series_python_filepath,
)

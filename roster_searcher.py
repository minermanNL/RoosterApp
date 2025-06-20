import pandas as pd
import requests
from datetime import datetime
import os
import re
from typing import List, Dict, Optional

class RosterSearcher:
    def __init__(self):
        self.workbook_data = {}
        self.search_results = []

    def read_excel_file(self, file_path: str, password: Optional[str] = None) -> Dict[str, pd.DataFrame]:
        try:
            if file_path.startswith(('http://', 'https://')):
                return self._read_from_url(file_path)
            elif os.path.exists(file_path):
                return self._read_local_file(file_path, password)
            else:
                raise FileNotFoundError(f"File not found: {file_path}")
        except Exception as e:
            print(f"Error reading Excel file: {str(e)}")
            return {}

    def _convert_sharepoint_url_to_download(self, url: str) -> str:
        if 'sharepoint.com/:x:/g/' in url:
            download_url = url.replace(':x:', ':u:').split('?')[0] + '?download=1'
            print(f"Converting SharePoint viewer URL to download URL: {download_url}")
            return download_url
        return url

    def _read_from_url(self, url: str) -> Dict[str, pd.DataFrame]:
        try:
            if 'sharepoint.com' in url:
                url = self._convert_sharepoint_url_to_download(url)
            print(f"Attempting to download from: {url}")
            response = requests.get(url)
            response.raise_for_status()
            temp_file = 'temp_excel.xlsx'
            with open(temp_file, 'wb') as f:
                f.write(response.content)
            with open(temp_file, 'rb') as f:
                sig = f.read(8)
            is_xlsx = sig.startswith(b'PK\x03\x04')
            is_xls = sig.startswith(b'\xD0\xCF\x11\xE0')
            if not (is_xlsx or is_xls):
                print("Downloaded file is not a valid Excel file. This may be an authentication page or error message.")
                if os.path.exists(temp_file):
                    os.remove(temp_file)
                if 'sharepoint.com' in url:
                    raise ValueError("SharePoint link format detected but couldn't download the Excel file directly. "
                                     "Please open the link in your browser, download the file, and then select it using Browse.")
                else:
                    raise ValueError("Downloaded file is not a valid Excel file. Please check if the link requires authentication and download manually if needed.")
            result = self._read_local_file(temp_file)
            if os.path.exists(temp_file):
                os.remove(temp_file)
            return result
        except requests.HTTPError as e:
            print(f"HTTP error reading from URL: {str(e)}")
            if ('sharepoint.com' in url or '1drv.ms' in url or 'onedrive.live.com' in url):
                print("SharePoint/OneDrive link may require authentication. Please open in browser and download manually.")
            return {}
        except Exception as e:
            print(f"Error reading from URL: {str(e)}")
            return {}

    def _read_local_file(self, file_path: str, password: Optional[str] = None) -> Dict[str, pd.DataFrame]:
        sheets_data = {}
        excel_file_obj = None
        engine_to_use = None
        if file_path.lower().endswith('.xlsx'):
            engine_to_use = 'openpyxl'
        try:
            if password:
                excel_file_obj = pd.ExcelFile(file_path, password=password, engine=engine_to_use)
            else:
                excel_file_obj = pd.ExcelFile(file_path, engine=engine_to_use)
            for sheet_name in excel_file_obj.sheet_names:
                try:
                    df = excel_file_obj.parse(sheet_name=sheet_name, header=None)
                    sheets_data[sheet_name] = df
                except Exception as e_sheet:
                    print(f"Warning: Could not read sheet '{sheet_name}': {str(e_sheet)}")
            return sheets_data
        except Exception as e_file:
            print(f"Error reading local file '{file_path}': {str(e_file)}")
            if "Excel file format cannot be determined" in str(e_file) or "engine" in str(e_file).lower():
                print("Pandas could not determine the Excel file format or the specified engine failed. "
                      "The file might be corrupted, not a standard Excel format, or an issue with the Excel engine (e.g., openpyxl). "
                      "If this is a SharePoint/OneDrive link, the downloaded file might not be the actual Excel data.")
            return {}
        finally:
            if excel_file_obj:
                try:
                    excel_file_obj.close()
                except Exception as e_close:
                    print(f"Warning: Error closing Excel file object: {str(e_close)}")

    def find_tables_in_sheet(self, df: pd.DataFrame) -> List[Dict]:
        tables = []
        rows, cols = df.shape
        for i in range(min(50, rows)):
            for j in range(min(20, cols)):
                cell_value = str(df.iloc[i, j]).strip() if pd.notna(df.iloc[i, j]) else ""
                if re.match(r'Week\s*\d+', cell_value, re.IGNORECASE):
                    table_info = self._extract_table_from_position(df, i, j)
                    if table_info:
                        tables.append(table_info)
                elif 'kandidaten' in cell_value.lower():
                    table_info = self._extract_kandidaten_table(df, i, j)
                    if table_info:
                        tables.append(table_info)
        # DEBUG: Print first 5 rows of each found table
        if tables:
            print(f"\nDEBUG: Preview of found tables in this sheet:")
            for idx, table in enumerate(tables):
                if 'data' in table:
                    print(f"\nTable {idx+1} (type: {table.get('type','?')}):")
                    print(table['data'].head(5))
                elif 'candidates' in table:
                    print(f"\nTable {idx+1} (type: kandidaten):")
                    for cand in table['candidates'][:5]:
                        print(cand)
        return tables

    def _extract_table_from_position(self, df: pd.DataFrame, start_row: int, start_col: int) -> Dict:
        try:
            max_row = start_row
            max_col = start_col
            for i in range(start_row, min(start_row + 50, len(df))):
                has_data = False
                for j in range(start_col, min(start_col + 20, len(df.columns))):
                    if pd.notna(df.iloc[i, j]) and str(df.iloc[i, j]).strip():
                        has_data = True
                        max_col = max(max_col, j)
                if has_data:
                    max_row = i
                else:
                    break
            if max_row > start_row and max_col > start_col:
                table_data = df.iloc[start_row:max_row+1, start_col:max_col+1].copy()
                return {
                    'type': 'schedule',
                    'start_row': start_row,
                    'start_col': start_col,
                    'data': table_data,
                    'header_row': 0
                }
        except Exception:
            pass
        return None

    def _extract_kandidaten_table(self, df: pd.DataFrame, start_row: int, start_col: int) -> Dict:
        try:
            candidates = []
            for i in range(start_row + 1, min(start_row + 100, len(df))):
                if pd.notna(df.iloc[i, start_col]):
                    candidate = str(df.iloc[i, start_col]).strip()
                    if candidate and not candidate.lower().startswith('week'):
                        candidates.append({
                            'number': i - start_row,
                            'name': candidate,
                            'row': i
                        })
                else:
                    break
            if candidates:
                return {
                    'type': 'kandidaten',
                    'start_row': start_row,
                    'start_col': start_col,
                    'candidates': candidates
                }
        except Exception:
            pass
        return None

    def search_name_in_tables(self, name: str, tables: List[Dict]) -> List[Dict]:
        results = []
        for table in tables:
            if table['type'] == 'schedule':
                dates = self._search_in_schedule_table(name, table)
                results.extend(dates)
            elif table['type'] == 'kandidaten':
                person_number = self._find_person_number(name, table)
                if person_number:
                    for other_table in tables:
                        if other_table['type'] == 'schedule':
                            dates = self._search_by_number_in_schedule(person_number, other_table)
                            results.extend(dates)
        return results

    def _find_person_number(self, name: str, kandidaten_table: Dict) -> Optional[int]:
        for candidate in kandidaten_table['candidates']:
            if name.lower() in candidate['name'].lower():
                return candidate['number']
        return None

    def _search_in_schedule_table(self, name: str, table: Dict) -> List[Dict]:
        results = []
        df = table['data']
        dates = self._get_schedule_dates(table)
        for i in range(1, len(df)):
            for col in dates:
                cell_value = str(df.iloc[i, col]) if pd.notna(df.iloc[i, col]) else ""
                if name.lower() in cell_value.lower():
                    date = dates.get(col)
                    if date:
                        results.append({
                            'name': name,
                            'date': date,
                            'position': f"Row {i+1}, Col {col+1}",
                            'context': cell_value,
                            'table_type': 'schedule'
                        })
        return results

    def _search_by_number_in_schedule(self, person_number: int, table: Dict) -> List[Dict]:
        results = []
        df = table['data']
        dates = self._get_schedule_dates(table)
        for i in range(1, len(df)):
            if pd.notna(df.iloc[i, 0]) and str(df.iloc[i, 0]).strip() == str(person_number):
                for col in dates:
                    cell_value = str(df.iloc[i, col]) if pd.notna(df.iloc[i, col]) else ""
                    if cell_value.strip():
                        date = dates.get(col)
                        if date:
                            results.append({
                                'name': f"Person #{person_number}",
                                'date': date,
                                'position': f"Row {i+1}, Col {col+1}",
                                'context': cell_value,
                                'table_type': 'schedule_by_number'
                            })
        return results

    def _extract_dates_from_table(self, df: pd.DataFrame) -> Dict:
        dates = {}
        for i in range(min(5, len(df))):
            for j in range(len(df.columns)):
                cell_value = str(df.iloc[i, j]) if pd.notna(df.iloc[i, j]) else ""
                date_patterns = [
                    r'\d{1,2}-\d{1,2}-\d{4}',
                    r'\d{1,2}/\d{1,2}/\d{4}',
                    r'\d{4}-\d{1,2}-\d{1,2}',
                    r'\d{1,2}\.\d{1,2}\.\d{4}'
                ]
                for pattern in date_patterns:
                    if re.search(pattern, cell_value):
                        dates[f"{i}_{j}"] = cell_value
                        dates[f"col_{j}"] = cell_value
                        break
        return dates

    def _get_schedule_dates(self, table: Dict) -> Dict[int, str]:
        df = table['data']
        header = df.iloc[0]
        week_cell = header.iloc[0]
        match = re.search(r'\d+', str(week_cell))
        if not match:
            return {}
        week_num = int(match.group())
        year = datetime.now().year
        day_map = {
            'maandag': 1, 'dinsdag': 2, 'woensdag': 3, 'donderdag': 4,
            'vrijdag': 5, 'zaterdag': 6, 'zondag': 7
        }
        dates = {}
        for col_idx, cell in enumerate(header):
            dayname = str(cell).strip().lower()
            if dayname in day_map:
                try:
                    dt = datetime.fromisocalendar(year, week_num, day_map[dayname])
                    dates[col_idx] = dt.strftime('%Y-%m-%d')
                except Exception:
                    pass
        return dates

    def search_person_schedule(self, file_path: str, person_name: str, password: Optional[str] = None) -> List[Dict]:
        """
        Main function to search for a person's schedule across all tables
        """
        print(f"Searching for '{person_name}' in {file_path}")
        
        # Read the Excel file
        sheets_data = self.read_excel_file(file_path, password)
        
        if not sheets_data:
            print("No data found in Excel file")
            return []
        
        all_results = []
        
        # Process each sheet
        for sheet_name, df in sheets_data.items():
            print(f"\nProcessing sheet: {sheet_name}")
            
            # Find tables in this sheet
            tables = self.find_tables_in_sheet(df)
            print(f"Found {len(tables)} tables in sheet '{sheet_name}'")
            
            # Search for the person in all tables
            results = self.search_name_in_tables(person_name, tables)
            
            # Add sheet info to results
            for result in results:
                result['sheet'] = sheet_name
            
            all_results.extend(results)
        
        return all_results

    def display_results(self, results: List[Dict]):
        """Display search results in a formatted way"""
        if not results:
            print("No matches found.")
            return
        
        print(f"\nFound {len(results)} work assignments:")
        print("-" * 80)
        
        # Group by sheet
        sheets = {}
        for result in results:
            sheet = result.get('sheet', 'Unknown')
            if sheet not in sheets:
                sheets[sheet] = []
            sheets[sheet].append(result)
        
        for sheet_name, sheet_results in sheets.items():
            print(f"\nSheet: {sheet_name}")
            print("-" * 40)
            
            for result in sheet_results:
                print(f"Date: {result.get('date', 'Unknown')}")
                print(f"Context: {result.get('context', '')}")
                print(f"Position: {result.get('position', '')}")
                print(f"Table Type: {result.get('table_type', '')}")
                print("-" * 30)

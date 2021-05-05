import json
import numpy as np
import pandas as pd
from xlsxwriter.utility import xl_cell_to_rowcol as xl
import googleapi.client


class SpreadSheet():
    """ A class for a spreadsheet object."""

    def __init__(self, client=None, response=None, **kwargs):

        if client is None:
            client = googleapi.client.Client()
        self.client = client

        if response is None:
            response = self.create(**kwargs)
        elif isinstance(response, str):
            response = self.get_workbook(id=response)

        if isinstance(response, dict):
            prop = response['properties']
            self._id = response['spreadsheetId']
            self._title = prop.get('title')
            self._locale = prop.get('locale')
            self._autoRecalc = prop.get('autoRecalc')
            self._timeZone = prop.get('timeZone')
            self._defaultFormat = prop.get('defaultFormat')
            self._spreadsheetTheme = prop.get('spreadsheetTheme')

            sheets_json = response['sheets']
            self._sheets = [Sheet(self, _) for _ in sheets_json]
        else:
            raise ValueError()

        self._current_datarange = None

    @property
    def id(self):
        """ID of the spreadsheet."""
        return self._id

    @property
    def title(self):
        """Title of the spreadsheet."""
        return self._title

    @property
    def sheet1(self):
        """Direct access to the first worksheet."""
        return self.get_sheet(0)

    @property
    def url(self):
        """URL of the spreadsheet."""
        return f'https://docs.google.com/spreadsheets/d/{self.id}'

    @property
    def properties(self):
        """Return the Google API formatted JSON of the properties
        for this SpreadSheet.
        """
        properties = {
            'properties': {
                'autoRecalc': self._autoRecalc,
                'defaultFormat': self._defaultFormat,
                'locale': self._locale,
                'spreadsheetTheme': self._spreadsheetTheme,
                'timeZone': self._timeZone,
                'title': self._title,
            },
            'sheets': [_.properties for _ in self._sheets],
            'spreadsheetId': self.id,
            'spreadsheetUrl': self.url,
        }

        return properties

    def create(
        self, title, sheet_title='Data',
        rows=1000, cols=26, freeze=None
    ):
        """Create a Spreadsheet."""
        body = {
            'properties': {
                'title': title
            },
            'sheets': [
                {
                    'properties': {
                        'title': sheet_title,
                        'gridProperties': {
                            'rowCount': rows,
                            'columnCount': cols,
                        }
                    }
                }
            ]
        }

        if isinstance(freeze, (list, tuple)):
            gp = body['sheets'][0]['properties']['gridProperties']
            gp['frozenRowCount'] = freeze[0]
            gp['frozenColumnCount'] = freeze[1]

        request = self.client.api['sheets'].spreadsheets().create(body=body)
        response = self.client._execute_requests(request)

        return response

    def get_workbook(self, id):
        """Retrive an existing workbook by ID"""
        request = self.client.api['sheets'].spreadsheets().get(
            spreadsheetId=id)
        response = self.client._execute_requests(request)

        return response

    def get_sheet(self, index):
        """Returns the worksheet with the specified index or title. Index by
        title first, and then by index.
        """
        sheets = self._sheets

        if not isinstance(index, (str, int)):
            ValueError('Specify integer index or title')

        if isinstance(index, str):
            sheets = [_ for _ in sheets if getattr(_, 'title') == index]
        if isinstance(index, int):
            sheets = [_ for _ in sheets if getattr(_, 'index') == index]

        if len(sheets) == 0:
            sheet = self.add_sheet(index)
        else:
            sheet = sheets[0]

        return sheet

    def add_sheet(self, title, rows=1000, cols=26, freeze=None):
        """ """
        request = {
            'addSheet': {
                'properties': {
                    'title': title,
                    'gridProperties': {
                        'rowCount': rows,
                        'columnCount': cols,
                    },
                }
            }
        }

        if isinstance(freeze, (list, tuple)):
            gp = request['addSheet']['properties']['gridProperties']
            gp['frozenRowCount'] = freeze[0]
            gp['frozenColumnCount'] = freeze[1]

        request = self.client.api['sheets'].spreadsheets().batchUpdate(
            spreadsheetId=self.id, body={'requests': request})
        response = self.client._execute_requests(request)

        nsheet = Sheet(self, response['replies'][0]['addSheet'])
        self._sheets += [nsheet]

        return nsheet

    def share(self, email, role='reader', message=None):
        """Share permissions, specific to an individual user."""
        body = {
            'kind': 'drive#permission',
            'type': 'user',
            'role': role,
            'emailAddress': email
        }

        if isinstance(message, str):
            body['emailMessage'] = message

        request = self.client.api['drive'].permissions().create(
            fileId=self.id, body=body)
        response = self.client._execute_requests(request)

        return response

    def __repr__(self):
        return json.dumps(self.properties, indent=4, separators=(',', ': '))


class Sheet():
    """A class for a worksheet object."""

    def __init__(self, spreadsheet, response):
        self._spreadsheet = spreadsheet
        self._title = response['properties'].get('title', '')
        self._sheetId = response['properties']['sheetId']
        self._index = response['properties']['index']
        self._sheetType = response['properties']['sheetType']
        self._hidden = response['properties'].get('hidden', None)
        self._tabColor = response['properties'].get('tabColor', {})
        self._rightToLeft = response['properties'].get('rightToLeft', False)

        grid_json = response['properties']['gridProperties']
        self.grid = Grid(response=grid_json)

    @property
    def id(self):
        """ID of the worksheet."""
        return self._sheetId

    @property
    def title(self):
        """Title of the worksheet."""
        return self._title

    @property
    def index(self):
        """Index of the worksheet."""
        return self._index

    @property
    def properties(self):
        """Return the Google API formatted JSON of the properties for this sheet.
        """
        properties = {
            'properties': {
                'title': self._title,
                'sheetId': self._sheetId,
                'index': self._index,
                'sheetType': self._sheetType,
                'hidden': self._hidden,
                'tabColor': self._tabColor,
                'rightToLeft': self._rightToLeft,
                'gridProperties': self.grid.properties,
            }
        }

        return properties

    def _add_sheet(self, request):
        """Check request for missing sheet names."""
        for i in request:
            if 'repeatCell' in i:
                i['repeatCell']['range']['sheetId'] = self.id
            elif 'updateSheetProperties' in i:
                i['updateSheetProperties']['properties']['sheetId'] = self.id
            elif 'autoResizeDimensions' in i:
                i['autoResizeDimensions']['dimensions']['sheetId'] = self.id
            elif 'updateDimensionProperties' in i:
                i['updateDimensionProperties']['range']['sheetId'] = self.id

        return request

    def update(self, request):
        """Perform and general update on a worksheet."""
        request = self._add_sheet(request)

        sh = self._spreadsheet
        request = sh.client.api['sheets'].spreadsheets().batchUpdate(
            spreadsheetId=sh.id, body={'requests': request})

        response = sh.client._execute_requests(request)
        return response

    def set_values(self, data, range='A1', valueInputOption='RAW'):
        """Update the values of a spreadsheet."""
        if isinstance(data, pd.DataFrame):
            _data = np.vstack([data.columns, data.values]).tolist()

        sh = self._spreadsheet
        request = sh.client.api['sheets'].spreadsheets().values().update(
            spreadsheetId=sh.id,
            range=f'{self.title}!{range}',
            valueInputOption=valueInputOption,
            body={'values': _data}
        )
        response = sh.client._execute_requests(request)

        data = DataRange(response, data, sheetId=self.id)
        self._spreadsheet._current_datarange = data

        return data

    def clear_values(self, rng):

        sh = self._spreadsheet
        request = sh.client.api['sheets'].spreadsheets().values().clear(
            spreadsheetId=sh.id,
            range=f'{self.title}!{rng}'
        )
        response = sh.client._execute_requests(request)

        return response

    def get_values(self, range):
        """Get the values of a spreadsheet."""
        sh = self._spreadsheet
        request = sh.client.api['sheets'].spreadsheets().values().get(
            spreadsheetId=sh.id,
            range=f'{self.title}!{range}'
        )
        response = sh.client._execute_requests(request)

        data = DataRange(response, sheetId=self.id)
        self._spreadsheet._current_datarange = data

        return data

    def add_pivot(
            self, rows, values, columns=None, filters=None,
            position='A1', datarange=None):
        """TODO """
        if datarange is None:
            datarange = self._spreadsheet._current_datarange

        if isinstance(position, str):
            position = xl(position)

        ro = []
        if isinstance(rows, list):
            for n in rows:
                d = {'sortOrder': 'ASCENDING', 'showTotals': True}
                d['sourceColumnOffset'] = datarange.get_loc(n)
                ro += [d]
        elif isinstance(rows, dict):
            for k, v in rows.items():
                d = {'sortOrder': 'ASCENDING', 'showTotals': False}
                d['sourceColumnOffset'] = datarange.get_loc(k)
                for k0, v0 in v.items():
                    d[k0] = v0
                ro += [d]

        co = []
        if isinstance(columns, list):
            for n in columns:
                d = {'sortOrder': 'ASCENDING', 'showTotals': False}
                d['sourceColumnOffset'] = datarange.get_loc(n)
                co += [d]
        elif isinstance(columns, dict):
            for k, v in columns.items():
                d = {'sortOrder': 'ASCENDING', 'showTotals': False}
                d['sourceColumnOffset'] = datarange.get_loc(k)
                for k0, v0 in v.items():
                    d[k0] = v0
                co += [d]

        fi = {}
        if filters is None:
            filters = {}
        else:
            for k, v in filters.items():
                if not isinstance(v, list):
                    v = [v]
                fi[datarange.get_loc(k)] = {'visibleValues': v}

        va = []
        for n in values:
            if isinstance(n, str):
                d = {'name': n}
                d['sourceColumnOffset'] = datarange.get_loc(n)
                d['summarizeFunction'] = 'SUM'
                va += [d]
            elif isinstance(n, dict):
                for k, v in n.items():
                    d = {'name': k}
                    if v.startswith('='):
                        d['summarizeFunction'] = 'CUSTOM'
                        d['formula'] = v
                    else:
                        d['sourceColumnOffset'] = datarange.get_loc(k)
                        d['summarizeFunction'] = v
                    va += [d]

        request = [{
            'updateCells': {
                'rows': {
                    'values': [
                        {
                            'pivotTable': {
                                'source': {
                                    'sheetId': datarange.sheetId,
                                    'startColumnIndex': datarange.startColumnIndex,
                                    'endColumnIndex': datarange.endColumnIndex,
                                    'startRowIndex': datarange.startRowIndex,
                                    'endRowIndex': datarange.endRowIndex
                                },
                                'rows': list(ro),
                                'columns': list(co),
                                'values': list(va),
                                'criteria': fi,
                                'valueLayout': 'HORIZONTAL'
                            }
                        }
                    ]
                },
                'start': {
                    'sheetId': self._sheetId,
                    'rowIndex': position[0],
                    'columnIndex': position[1]
                },
                'fields': 'pivotTable'
            },
        }]

        response = self.update(request)
        return response

    def add_slicer(
        self, tag, position='A1', filter=None,
        title=None, datarange=None
    ):
        """ """
        if datarange is None:
            datarange = self._spreadsheet._current_datarange

        if not isinstance(filter, dict):
            filter = {}

        if title is None:
            title = tag

        if isinstance(position, str):
            position = xl(position)

        slicer = {
            'spec': {
                'dataRange': {
                    'sheetId': datarange.sheetId,
                    'startColumnIndex': datarange.startColumnIndex,
                    'endColumnIndex': datarange.endColumnIndex,
                    'startRowIndex': datarange.startRowIndex,
                    'endRowIndex': datarange.endRowIndex
                },
                'columnIndex': datarange.get_loc(tag),
                'applyToPivotTables': True,
                'title': title,
                'filterCriteria': filter
            },
            'position': {
                'overlayPosition': {
                    'anchorCell': {
                        'sheetId': self._sheetId,
                        'rowIndex': position[0],
                        'columnIndex': position[1],
                    }
                },
            },
        }

        request = [{'addSlicer': {'slicer': slicer}}]
        response = self.update(request)
        return response

    def __repr__(self):
        return json.dumps(self.properties, indent=4, separators=(',', ': '))


class Grid():
    """A Class for a Grid within a sheet."""

    def __init__(self, response):
        self._rowCount = response.get('rowCount')
        self._columnCount = response.get('columnCount')

        self._frozenRowCount = response.get('frozenRowCount')
        self._frozenColumnCount = response.get('frozenColumnCount')

        self._hideGridlines = response.get('hideGridlines')
        self._rowGroupControlAfter = response.get('rowGroupControlAfter')
        self._columnGroupControlAfter = response.get('columnGroupControlAfter')

    @property
    def properties(self):
        """Return the Google API formatted JSON of the properties
        for this Grid.
        """
        properties = {
            'columnCount': self._columnCount,
            'rowCount': self._rowCount,
            'frozenRowCount': self._frozenRowCount,
            'frozenColumnCount': self._frozenColumnCount,
            'hideGridlines': self._hideGridlines,
            'rowGroupControlAfter': self._rowGroupControlAfter,
            'columnGroupControlAfter': self._columnGroupControlAfter,
        }

        return properties


class DataRange():
    """ """

    def __init__(self, response, data=None, sheetId=None):
        self._spreadsheetId = response.get('spreadsheetId')

        rng = response.get('updatedRange', response.get('range')).split('!')
        if data is None:
            data = response.get('values')

        if isinstance(data, type(None)):
            raise ValueError('Data cannot be None')
        if not isinstance(data, pd.DataFrame):
            data = pd.DataFrame(
                data[1:], columns=[_.strip() for _ in data[0]])

        # Data Massage (May cause problems)
        data = data.replace(
            '[\$,)]', '', regex=True).replace('[(]', '-', regex=True)
        for col in data.columns:
            try:
                data[col] = pd.to_numeric(data[col])
                data[col].fillna(0, inplace=True)
            except ValueError:
                data[col] = pd.to_datetime(data[col], errors='ignore')

        self._data = data

        self._sheetId = sheetId
        self._sheetTitle = None
        if self._sheetId is None:
            self._sheetTitle = rng[0]

        ind = rng[1].split(':')[0]
        if not ind[-1].isdigit():
            ind += '1'
        ind = xl(ind)

        self._startIndex = ind
        self._endIndex = (
            ind[0] + self._data.shape[0],
            ind[1] + self._data.shape[1] - 1,
        )

    @property
    def spreadsheetId(self):
        """ """
        return self._spreadsheetId

    @property
    def sheetId(self):
        """ """
        return self._sheetId

    @property
    def startIndex(self):
        """ """
        return self._startIndex

    @property
    def endIndex(self):
        """ """
        return self._endIndex

    @property
    def data(self):
        """ """
        return self._data

    def get_loc(self, name):
        """ """
        return self._data.columns.get_loc(name)

    def __repr__(self):
        return str(self._data)

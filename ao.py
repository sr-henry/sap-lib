import win32com.client as win32

# Excel 1-* Workbooks 1-* DataSources


class SapAnalysisOfficeEx(Exception):
    def __init__(self, message):
        super().__init__(message)


class SapAnalysisOffice:

    def __init__(self, workbook_path):

        self.xl = None
        self.wb = None

        try:
            self.xl = win32.Dispatch('Excel.Application')
            self.wb = self.xl.Workbooks.Open(workbook_path, False, False)
        except win32.pywintypes.com_error as e:
            self.close()
            raise SapAnalysisOfficeEx(
                f'Excel | Workbook Dispatch exception occurred: {e}'
            )

        self.xl.DisplayAlerts = False
        self.xl.ScreenUpdating = False

        if not self.__activate_analysis_addin():
            self.close()
            raise SapAnalysisOfficeEx(
                'SAP Analysis for Office addin not found'
            )

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def __activate_analysis_addin(self) -> bool:
        for addin in self.xl.Application.COMAddIns:
            if addin.progID == 'SapExcelAddIn':
                addin.Connect = False
                addin.Connect = True
                return True
        return False

    @property
    def datasources(self):
        return [ds for ds, _ in self.xl.Application.Run('SAPListOf', 'DATASOURCES')]

    def logon(self, client: str, usr: str, pwd: str, ds: str) -> bool:
        return bool(
            self.xl.Run("SAPLogon", ds, *
                        [c for c in [client, usr, pwd] if c])
        )

    def is_connected(self, ds: str) -> bool:
        return bool(self.xl.Application.Run('SAPGetProperty', 'IsConnected', ds))

    def pause_variable_submit(self, state: bool) -> bool:
        return bool(self.xl.Run('SAPExecuteCommand', 'PauseVariableSubmit',
                    'On' if state else 'Off'))

    def set_refresh_behaviour(self, state: bool) -> bool:
        return bool(self.xl.Run('SAPSetRefreshBehaviour', 'On' if state else 'Off'))

    def refresh(self, ds: str) -> bool:
        return bool(self.xl.Run('SAPExecuteCommand', 'RefreshData', ds))

    def set_variable(self, tech_name, value, type, ds):
        return bool(self.xl.Run('SAPSetVariable', tech_name, value, type, ds))

    def close(self):
        if isinstance(self.xl, win32.CDispatch):
            if isinstance(self.wb, win32.CDispatch):
                self.wb.Close(True)
            self.xl.Quit()
            self.xl = None

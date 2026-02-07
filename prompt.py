import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application")
print([acc.DisplayName for acc in outlook.Session.Accounts])

import xlwings as xw
import openpyxl
from datetime import datetime as DT
import pandas as pd
import win32com.client
import os
import re



Assignee = {
    "AY" : "Aditya Yadav",
    "OS" : "Omkar Sagavekar",
    "MH" : "Masroor Hafiz",
}
main_file = "samples.xlsx"


def Outbound_auto():
    # Read the data
    df_outbound = pd.read_excel(main_file, sheet_name="Outbond failures", skiprows=5)
    DBname_LegalCompanyName_PartnerCode = (
        df_outbound[
            (df_outbound["Assignee"] == Assignee.get("AY"))
            & (df_outbound["Status"] == "Pending")
            & (
                df_outbound["Error"].str.contains(
                    "check with supplier", case=False, regex=True
                )
            )
        ][["DBname_LegalCompanyName_PartnerCode"]][
            "DBname_LegalCompanyName_PartnerCode"
        ]
        .unique()
        .tolist()
    )
    # print(len(DBname_LegalCompanyName_PartnerCode),DBname_LegalCompanyName_PartnerCode)

    wb = xw.Book(main_file)
    Outbond_failures = wb.sheets["Outbond failures"]

    def email(P, target_wb_path, td_in_draft):
        ol = win32com.client.Dispatch("outlook.application")
        olmailitem = 0x0  # size of the new email
        newmail = ol.CreateItem(olmailitem)
        newmail.Subject = f"{P}_PO document failure GEP/{P.split('_')[0]}| Reporting_date {DT.now().strftime('%m.%d.%Y')}"
        try:
            df_contact = pd.read_excel(main_file, sheet_name="Contact")

            newmail.To = df_contact[
                (df_contact["DBname_LegalCompanyName_PartnerCode"] == P)
            ][["OutboundEmail_To"]].to_string(index=False, header=False)

            newmail.CC = df_contact[
                (df_contact["DBname_LegalCompanyName_PartnerCode"] == P)
            ][["OutboundEmail_CC"]].to_string(index=False, header=False)
        except:
            newmail.To = "Add email address"
            newmail.CC = "Add email address"

        # Email draft in HTML formate
        newmail.HTMLBody = f"""
        <html>
        Hi {str(P.split('_')[1])} team, <br><br>
            
        Please refer attached speadsheet of Purchase order failures.<br><br>
        We are not able to submit those via integration due to response error.<br>

        Details are mentioned in the file. Please check and confirm the status or amendments.<br><br>
        
        <b>Customer name</b> : <u> {P.split('_')[0]} </u><br>
        <b>Vendor/supplier name</b> : {P.split('_')[1]} <br>
        <b>Integration type</b> : cXML/EDI  <br>
        <b>Reporting date</b>: {DT.now().strftime('%m.%d.%Y')} <br><br>

        {td_in_draft.to_html(index=False)}

        <br>
        
        Thanks and Regards,<br>
        Supplier Enablement Team_GEP 
        </html>
        """
        attach = target_wb_path
        newmail.Attachments.Add(attach)
        newmail.Display()  # --> To display the mail before sending it
        # newmail.Send()

    for P in DBname_LegalCompanyName_PartnerCode:
        # print(P)
        book = openpyxl.Workbook()
        # sheet = book.create_sheet(f'{P}')
        # supplier_worksheet.append(sheet)

        headers = [
            "DBname",
            "Error",
            "LegalCompanyName",
            "DocumentNumber",
            "OrderAmount",
            "DateModified",
            "PartnerCode",
            "DBname_LegalCompanyName_PartnerCode",
        ]
        book.active.append(headers)

        # supplier_workbook.append(book)
        book.save(f"PO_{P}_{DT.now().strftime('%m.%d.%Y')}.xlsx")
        xw.Book(f"PO_{P}_{DT.now().strftime('%m.%d.%Y')}.xlsx").sheets["Sheet"].range(
            "A1:K1"
        ).columns.autofit()
        xw.Book(f"PO_{P}_{DT.now().strftime('%m.%d.%Y')}.xlsx").sheets["Sheet"].range(
            "A1:K1"
        ).api.Font.Bold = True
        xw.Book(f"PO_{P}_{DT.now().strftime('%m.%d.%Y')}.xlsx").save()
        xw.Book(f"PO_{P}_{DT.now().strftime('%m.%d.%Y')}.xlsx").close()

    for P in DBname_LegalCompanyName_PartnerCode:
        data_range = (
            Outbond_failures.range("A6:X7")
            .expand()
            .options(pd.DataFrame, index=False, header=True)
            .value
        )

        filtered_data = data_range[
            (data_range["Assignee"] == Assignee.get("AY"))  # Condition 1
            & (data_range["Status"] == "Pending")  # Condition 2
            & (
                data_range["Error"].str.contains(
                    "check with supplier", case=False, regex=True
                )
            )  # Condition 3
            & (data_range["DBname_LegalCompanyName_PartnerCode"] == P)  # Condition 4
        ]

        selected_columns = filtered_data[
            [
                "DBname",
                "Error",
                "LegalCompanyName",
                "DocumentNumber",
                "OrderAmount",
                "DateModified",
                "PartnerCode",
                "DBname_LegalCompanyName_PartnerCode",
            ]
        ]

        td_in_draft = filtered_data[
            ["DocumentNumber", "OrderAmount", "Error", "DateModified"]
        ]
        # Open the target workbook
        target_wb = xw.Book(f"PO_{P}_{DT.now().strftime('%m.%d.%Y')}.xlsx")
        target_ws = target_wb.sheets.active

        # Write the filtered data from the DataFrame to the target worksheet
        target_ws.range("A1").options(index=False, header=True).value = selected_columns

        # Autofit the "DateModified" column
        # Assuming "DateModified" is in column F
        target_ws.range("F:F").autofit()
        target_ws.range("B:B").autofit()  # Error column

        # Disable text wrapping for the specified named range
        target_ws.range("A:I").api.WrapText = False

        # Save and close the target workbook
        target_wb.save()
        target_wb_path = os.path.abspath(target_wb.fullname)
        target_wb.close()

        # calling email function
        email(P, target_wb_path, td_in_draft)

    # Changing the status column from original file
    for row in Outbond_failures.used_range.rows[6:]:
        if (
            row[3].value == Assignee.get("AY")
            and re.search("check with supplier", str(row[4].value), re.IGNORECASE)
            and row[5].value == "Pending"
        ):
            row[5].value = "Sending to supplier failed"
            row[6].value = "As per Resolution Column - Email sent to supplier to check"
            row[7].value = DT.now().strftime("%m/%d/%Y")


Outbound_auto()

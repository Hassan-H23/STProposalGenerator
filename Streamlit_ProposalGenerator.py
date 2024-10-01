from datetime import datetime
from io import BytesIO
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches
from docx.shared import Pt
import pandas as pd
import streamlit as st
import Service


def set_cell_margins(cell, top, right, bottom, left):
    cell._element.get_or_add_tcPr().append(
        parse_xml(
            r'<w:tcMar {}>'
            r'<w:top w:w="{top}" w:type="dxa"/>'
            r'<w:right w:w="{right}" w:type="dxa"/>'
            r'<w:bottom w:w="{bottom}" w:type="dxa"/>'
            r'<w:left w:w="{left}" w:type="dxa"/>'
            r'</w:tcMar>'.format(nsdecls('w'), top=top, right=right, bottom=bottom, left=left)
        )
    )

# Initialize services_list in session state if it doesn't exist
if 'services_list' not in st.session_state:
    st.session_state.services_list = []
st.title('Maverick :blue[Proposals] :page_facing_up:')

# Get the path to the proposal template from the local machine
st.header('Provide path to proposal template', divider="gray")
doc_path = st.text_input("Template Path",
                         r"C:\Users\Usuario\PycharmProjects\WordTest\Standard Template .docx")

# Function to create a new service form
with st.sidebar:
    st.header('Maverick :blue[Proposals]', divider="gray")
    companyName = st.text_input("Company Name")
    ceoName = st.text_input("CEO Name")
    companyAddress = st.text_input("Company Address")

    with st.form(key=f"services_intake_form_{len(st.session_state.services_list)}", clear_on_submit=False):
        serviceChoice = st.selectbox("Service", ["Concierge", "Security", "Janitorial", "Cleaning", "Porter", "Valet"])
        weeklyHours = st.number_input("Weekly Hours", min_value=0, max_value=1000, step=1, value=168)
        billRate = st.number_input("Bill Rate", min_value=0.0, max_value=1000.0, step=1.0, value=27.0)
        yearlyHolidayHours = st.number_input("Yearly Holiday Hours", min_value=0, max_value=1000, step=1)
        inflationRate = st.number_input("Inflation Rate", min_value=0.0, max_value=100.0, step=0.1, value=3.0)
        submitted = st.form_submit_button("Add Service")
        finalSubmit = st.form_submit_button("Generate Proposal")

    if submitted:
        new_service = Service.Service(serviceChoice, weeklyHours, billRate, yearlyHolidayHours, inflationRate)
        st.session_state.services_list.append(new_service)
        st.rerun()  # Rerun the script to update the state

    if finalSubmit:
        # Legend
        current_date = datetime.now()


        # Create a new Document
        if doc_path:
            document = Document(doc_path)
        # Replace placeholders in the document if you are loading a template
        for paragraph in document.paragraphs:
            if 'CCNN' in paragraph.text:
                paragraph.text = paragraph.text.replace('CCNN', companyName)
            if 'NNNN' in paragraph.text:
                paragraph.text = paragraph.text.replace('NNNN', ceoName)
            if 'DDDD' in paragraph.text:
                paragraph.text = paragraph.text.replace('DDDD', current_date.strftime("%Y-%m-%d"))
            if 'CCAA' in paragraph.text:
                paragraph.text = paragraph.text.replace('CCAA', companyAddress)

        # Header
        document.add_paragraph()
        heading = document.add_heading('COST PROPOSAL', level=1)
        run = heading.runs[0]
        run.bold = True
        run.font.size = Pt(22)
        document.add_paragraph()

        # Table
        table = document.add_table(rows=1, cols=7)
        table.style = 'Table Grid'
        # Table Headers
        hdr_cells = table.rows[0].cells
        table_headers = ["Service", "Weekly Hours", "Bill Rate", "Monthly Amount",
                         f"Annual Amount (Year 1)", "Annual Amount (Year 2)",
                         "Annual Amount (Year 3)"]

        for i, header in enumerate(table_headers):
            hdr_cells[i].text = table_headers[i]
            hdr_cells[i].paragraphs[0].runs[0].font.bold = True
            hdr_cells[i].paragraphs[0].runs[0].font.size = Inches(0.15)
            hdr_cells[i].paragraphs[0].paragraph_format.alignment = 1
            set_cell_margins(hdr_cells[i], 55, 50, 45, 50)
            if i != 0:
                hdr_cells[i]._element.get_or_add_tcPr().append(
                    parse_xml(r'<w:shd {} w:fill="ADD8E6"/>'.format(nsdecls('w'))))
            else:
                hdr_cells[i]._element.get_or_add_tcPr().append(
                    parse_xml(r'<w:shd {} w:fill="000080"/>'.format(nsdecls('w'))))

        # Populate the table with service data
        for service in st.session_state.services_list:
            row_cells = table.add_row().cells
            row_cells[0].text = service.serviceName
            row_cells[1].text = str(service.weeklyHours)
            row_cells[2].text = f"${service.billRate:.2f}"
            row_cells[3].text = f"${service.monthlyAmount:.2f}"
            row_cells[4].text = f"${service.annualAmountYear1:.2f}"
            row_cells[5].text = f"${service.annualAmountYear2:.2f}"
            row_cells[6].text = f"${service.annualAmountYear3:.2f}"

            for cell in row_cells:
                cell.paragraphs[0].paragraph_format.alignment = 1
                set_cell_margins(cell, 100, 100, 100, 100)

        # Footer
        document.add_paragraph()
        document.add_paragraph("***Applicable NJ Sales Tax Included***").paragraph_format.alignment = 1
        document.add_paragraph(
            "***New Year’s Day, Presidents Day, Memorial Day, Independence Day, Labor Day, Thanksgiving Day, Christmas Day Is Included In the Above Pricing***")

        # Save the document to a BytesIO stream
        doc_io = BytesIO()
        document.save(doc_io)
        doc_io.seek(0)
        # Create a download button for the document
        if st.session_state.services_list:
            st.download_button(
                label="Download Proposal",
                data=doc_io,
                file_name=f"{companyName} Proposal.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

# Convert the list of services to a DataFrame
if st.session_state.services_list:
    services_df = pd.DataFrame([vars(service) for service in st.session_state.services_list])
    st.dataframe(services_df)

    # Allow user to remove or update services
    st.header('Update or Remove Service :pencil:', divider="gray")
    service_to_edit = st.selectbox("", options=[service.serviceName for service in st.session_state.services_list] + [
        "None"])

    if service_to_edit != "None":
        # Get selected service
        selected_service = next(
            service for service in st.session_state.services_list if service.serviceName == service_to_edit)

        with st.form(key='update_service_form'):
            updated_service_choice = st.selectbox("Service",
                                                  ["Concierge", "Security", "Janitorial", "Cleaning", "Porter",
                                                   "Valet"],
                                                  index=["Concierge", "Security", "Janitorial", "Cleaning", "Porter",
                                                         "Valet"].index(selected_service.serviceName))
            updated_weekly_hours = st.number_input("Weekly Hours", min_value=0, max_value=1000, step=1,
                                                   value=selected_service.weeklyHours)
            updated_bill_rate = st.number_input("Bill Rate", min_value=0.0, max_value=1000.0, step=1.0,
                                                value=selected_service.billRate)
            updated_yearly_holiday_hours = st.number_input("Yearly Holiday Hours", min_value=0, max_value=1000, step=1,
                                                           value=selected_service.yearlyHolidayHours)
            col1, col2, col3 = st.columns(3, vertical_alignment="bottom")
            with col1:
                update_button = st.form_submit_button("Update Service")
            with col2:
                remove_button = st.form_submit_button("Remove Service")
            with col3:
                clear_button = st.form_submit_button("Clear Services")

        if update_button:
            # Update the service with the new values
            selected_service.serviceName = updated_service_choice
            selected_service.weeklyHours = updated_weekly_hours
            selected_service.billRate = updated_bill_rate
            selected_service.yearlyHolidayHours = updated_yearly_holiday_hours
            st.success("Service updated successfully!")
            st.rerun()  # Rerun to update the state

        if remove_button:
            # Remove the service from the list
            st.session_state.services_list.remove(selected_service)
            st.success("Service removed successfully!")
            st.rerun()  # Rerun to update the state

        if clear_button:
            st.session_state.services_list.clear()
            st.rerun()  # Rerun to update the state
else:
    st.write("No services added")

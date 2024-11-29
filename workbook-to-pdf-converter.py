import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


def convert_excel_to_pdf(excel_file, output_pdf):
    # Read the Excel file
    xls = pd.ExcelFile(excel_file)

    # Create a PdfPages object to save multiple pages into the same PDF
    with PdfPages(output_pdf) as pdf:
        for sheet_name in xls.sheet_names:
            # Read the sheet into a DataFrame
            df = pd.read_excel(xls, sheet_name=sheet_name)

            # Plot the DataFrame
            fig, ax = plt.subplots(figsize=(8, 6))
            ax.axis('tight')
            ax.axis('off')

            # Create a table and plot it
            table = ax.table(cellText=df.values, colLabels=df.columns, loc='center', cellLoc='center',
                             bbox=[0, 0, 1, 1])

            # Customize table appearance (optional)
            table.auto_set_font_size(False)
            table.set_fontsize(8)
            table.auto_set_column_width(col=list(range(len(df.columns))))

            # Add the table to the PDF
            pdf.savefig(fig)
            plt.close(fig)

    print(f"PDF saved as {output_pdf}")


def main():

    convert_excel_to_pdf('input.xlsx', 'output.pdf')


if __name__ == '__main__':
    main()


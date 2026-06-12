"""ReportBuilder class that orchestrates Excel and Word report creation."""

from excel_report import create_excel_report
from word_report import create_word_report


class ReportBuilder:
    def __init__(self, df, df_hops, target_building, output_dir):
        self.df = df
        self.df_hops = df_hops
        self.target_building = target_building
        self.output_dir = output_dir

    def build_excel(self):
        """Create the Excel report and return the output path."""
        return create_excel_report(
            self.df,
            self.df_hops,
            self.target_building,
            self.output_dir,
        )

    def build_word(self, excel_path=None):
        """Create the Word report and return the output path."""
        return create_word_report(
            self.df,
            self.df_hops,
            self.target_building,
            self.output_dir,
            excel_path,
        )

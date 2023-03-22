from xlwt import easyxf
from main import arial10

PT = 20

style_string = "font: bold on; pattern: pattern solid, fore_color yellow; borders: top_color black, bottom_color black, right_color black, left_color black,left medium, right medium, top medium, bottom medium;"
sub_header_style = "font: bold on; pattern: pattern solid, fore_color yellow; borders: top_color black, bottom_color black, right_color black, left_color black,left medium, right medium, top medium, bottom medium;align: wrap yes, vert centre, horiz center;"
yellow_background_header = easyxf(style_string)
yellow_bg_sub_header = easyxf(sub_header_style)
bold_center = easyxf("align: wrap yes, vert centre, horiz center;font: name Arial,height 280, bold True")


class FitSheetWrapper(object):
    """Try to fit columns to max size of any entry.
    To use, wrap this around a worksheet returned from the
    workbook's add_sheet method, like follows:

        sheet = FitSheetWrapper(book.add_sheet(sheet_name))

    The worksheet interface remains the same: this is a drop-in wrapper
    for auto-sizing columns.
    """
    def __init__(self, sheet):
        self.sheet = sheet
        self.widths = dict()

    def write(self, r, c, label='', *args, **kwargs):
        self.sheet.write(r, c, label, *args, **kwargs)
        width = int(arial10.fitwidth(label))
        if width > self.widths.get(c, 0):
            self.widths[c] = width
            self.sheet.col(c).width = width

    def __getattr__(self, attr):
        return getattr(self.sheet, attr)

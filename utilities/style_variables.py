from pptx.enum.dml import MSO_THEME_COLOR_INDEX


def get_brand_color(color):
    colors = {
        '*WHITE': MSO_THEME_COLOR_INDEX.BACKGROUND_1,
        '*BLACK': MSO_THEME_COLOR_INDEX.TEXT_1,
        '*GRAY': MSO_THEME_COLOR_INDEX.BACKGROUND_2,
        '*BLUE': MSO_THEME_COLOR_INDEX.ACCENT_1,
        '*RED': MSO_THEME_COLOR_INDEX.ACCENT_2,
        '*GREEN': MSO_THEME_COLOR_INDEX.ACCENT_3,
        '*ORANGE': MSO_THEME_COLOR_INDEX.ACCENT_4,
        '*YELLOW': MSO_THEME_COLOR_INDEX.ACCENT_5,
        '*PURPLE': MSO_THEME_COLOR_INDEX.ACCENT_6
    }
    return colors.get(f"{'*' if '*' not in color else ''}{color.upper()}")


def get_chart_style(color='*BLUE'):
    chart_styles = {'*BLUE': 3, '*RED': 4, '*GREEN': 5, '*ORANGE': 6, '*YELLOW': 7, '*PURPLE': 8, '*MULTI': None}
    return chart_styles.get(color)

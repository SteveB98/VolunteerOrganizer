from openpyxl.styles import Font, Alignment
"""
Program: SheetFormat.py
Purpose: Contains OpenPYXL fon & alignment varialbes for ScheduleOrganizer.py
"""
title_font = Font(name='Calibri',
                size=28,
                bold=True,
                italic=False,
                underline='none',
                strike=False,
                color='FF000000')

bolded_font = Font(name='Calibri',
                size=12,
                bold=True,
                italic=False,
                underline='none',
                strike=False,
                color='FF000000')

supervisor_font = Font(name='Calibri',
                size=14,
                bold=True,
                italic=False,
                underline='none',
                strike=False,
                color='FF000000')

align = Alignment(horizontal='center',
                    vertical='bottom',
                    text_rotation=0,
                    wrap_text=True,
                    shrink_to_fit=False,
                    indent=0)
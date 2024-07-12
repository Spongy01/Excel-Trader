def render_single_component(sheet, start_range, end_range, data, color=(255, 255, 255), align_center=False,
                            merge=False):
    if merge:
        sheet.range(start_range, end_range).merge()
    if align_center:
        sheet.range(start_range, end_range).api.HorizontalAlignment = -4108
    sheet.range(start_range, end_range).color = color
    sheet.range(start_range, end_range).value = data
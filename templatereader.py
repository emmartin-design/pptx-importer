from pptx import Presentation
import logging
import variables as v



# Indicates preferred status for layouts with one body field and charts of equal size
def preferred_layouts(ld, type):
    if type == 'scorecard':
        for idx in ld:
            if ld[idx]['name'] in v.scorecard_layouts:
                ld[idx]['preferred'] = True
    elif type == 'dtv':
        for idx in ld:
            if ld[idx]['name'] in v.dtv_layouts:
                ld[idx]['preferred'] = True
    else:
        for idx in ld:
            if ld[idx]['bodycount'] == 2:  # One for the section tag, one for body copy
                for cht in ['chartcount', 'tablecount']:
                    if ld[idx][cht] == 1:
                        for ph in ld[idx]:
                            for x in ['CHART', 'TABLE']:
                                if x in ph:
                                    if ld[idx][ph]['width'] == 9.12:
                                        ld[idx]['preferred'] = True
                    elif ld[idx][cht] > 1:
                        if len(set(ld[idx]['widthtest'])) == 1:
                            if max(ld[idx]['widthtest']) < 9.12:
                                ld[idx]['preferred'] = True
            elif ld[idx]['bodycount'] == 1:  # Only for pages with 4 charts
                if ld[idx]['chartcount'] == 4:
                    if max(ld[idx]['widthtest']) > 3:
                        ld[idx]['preferred'] = True


    ld_pared = ld.copy()  # Copies updated ld for paring to just essentials
    for idx in ld:
        if ld[idx]['preferred'] == False:  # Checks to see if slide is good for import
            del ld_pared[idx]  # If it is not, it deletes it from the copy.
    return ld_pared


def readtemplate(template, type):  # This collects information about the template and drops it into a dictionary
    layout_dict = {}
    scorecard_dict = {}
    prs = Presentation(template)
    for idx, layout in enumerate(prs.slide_layouts):
        layout_dict[idx] = {}
        chartcount, tablecount, bodycount, titlecount, picturecount = 0, 0, 0, 0, 0
        widths = []  # Tracks width of all charts
        for shidx, shape in enumerate(layout.placeholders):
            x = shape.placeholder_format.idx
            stype = str(shape.placeholder_format.type)
            shidxstr = str(shidx)
            for shapetype in v.phtypes:
                if shapetype in stype:
                    shapename = shapetype + ' ' + shidxstr
                    layout_dict[idx][shapename] = {
                        'index': x,
                        'width': round(shape.width.inches, 2),
                        'height': round(shape.height.inches, 2),
                        'leftloc': round(shape.left.inches, 2)
                    }
                    if 'CHART' in shapetype:
                        chartcount += 1
                        widths.append(round(shape.width.inches, 2))
                    elif 'TABLE' in shapetype:
                        tablecount += 1
                        widths.append(round(shape.width.inches, 2))
                    elif shapetype == 'BODY':
                        bodycount += 1
                    elif 'TITLE' in shapetype:
                        titlecount += 1
                    elif 'PICTURE' in shapetype:
                        widths.append(round(shape.width.inches, 2))
                        picturecount += 1

        layout_dict[idx]['name'] = layout.name
        layout_dict[idx]['chartcount'] = chartcount
        layout_dict[idx]['tablecount'] = tablecount
        layout_dict[idx]['bodycount'] = bodycount
        layout_dict[idx]['titlecount'] = titlecount
        layout_dict[idx]['picturecount'] = picturecount
        layout_dict[idx]['widthtest'] = widths
        layout_dict[idx]['preferred'] = False


    ld = preferred_layouts(layout_dict, type)

    logging.info('Template analyzed')
    return ld

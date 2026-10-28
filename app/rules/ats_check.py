from app.schema.template import ATS_SAFE_FONTS, STANDARD_HEADINGS, Template


def ats_issues(t: Template) -> list[str]:
    a, issues = t.ats, []
    if a.mode == "off":
        return issues
    if a.single_column and t.page.columns > 1:
        issues.append("The page uses more than one column.")
    if a.no_tables and t.structure.uses_tables:
        issues.append("Content sits inside tables; some parsers read them out of order.")
    if a.no_text_boxes and t.structure.uses_text_boxes:
        issues.append("Content sits inside text boxes; some parsers skip them.")
    if a.no_images and (t.structure.has_images or t.header.show_photo):
        issues.append("The file has images or a photo.")
    if a.contact_in_body and t.structure.contact_in_page_header:
        issues.append("Contact details are in the Word header; some parsers skip headers.")
    fonts = {f for f in (t.typography.primary_font, t.typography.secondary_font) if f}
    if a.safe_fonts_only and not fonts <= ATS_SAFE_FONTS:
        issues.append(f"Fonts outside the safe list: {sorted(fonts - ATS_SAFE_FONTS)}")
    body = t.typography.body.size_pt
    if body and not a.min_body_pt <= body <= a.max_body_pt:
        issues.append(f"Body text is {body} pt (recommended {a.min_body_pt}-{a.max_body_pt}).")
    margins = [m for m in (t.page.margin_top_in, t.page.margin_bottom_in,
                           t.page.margin_left_in, t.page.margin_right_in) if m is not None]
    if margins and min(margins) < a.min_margin_in:
        issues.append(f"A margin is below {a.min_margin_in} in.")
    if t.links.display == "label":
        issues.append("Links show only a label, so a parser may lose the URL.")
    if a.standard_headings:
        for s in t.sections:
            known = STANDARD_HEADINGS.get(s.key)
            if s.visible and known and s.heading.strip().lower() not in known:
                issues.append(f"Heading '{s.heading}' is not a standard name for {s.key}.")
    return issues
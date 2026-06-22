import pikepdf

def write_bookmarks_to_pdf(input_pdf_path: str, output_pdf_path: str, plan: list):
    """
    Reads the original PDF, applies the plan, and saves the new PDF.
    """
    with pikepdf.open(input_pdf_path) as pdf:
        with pdf.open_outline() as outline:
            outline.root.clear()
            for group in plan:
                items = group["items"]
                if not items:
                    continue
                # Level 1: Visit / Form
                group_item = pikepdf.models.outlines.OutlineItem(
                    title       = group["group"],
                    destination = items[0]["page"] - 1,
                )
                for item in items:
                    # Level 2: visit name or form name
                    level2 = pikepdf.models.outlines.OutlineItem(
                        title       = item["name"],
                        destination = item["page"] - 1,
                    )
                    # Level 3: form names (under visit) or visit names (under form)
                    for child in item.get("children", []):
                        level3 = pikepdf.models.outlines.OutlineItem(
                            title       = child["name"],
                            destination = child["page"] - 1,
                        )
                        level2.children.append(level3)
                    group_item.children.append(level2)
                outline.root.append(group_item)
                
        pdf.save(output_pdf_path)

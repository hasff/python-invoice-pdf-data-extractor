import fitz  # pymupdf


def draw_vertical_lines(input_pdf, output_pdf, x_positions, color=(1, 0, 0)):
    doc = fitz.open(input_pdf)

    for page in doc:
        height = page.rect.height

        for x in x_positions:

            p1 = fitz.Point(x, 0)
            p2 = fitz.Point(x, height)

            page.draw_line(
                p1,
                p2,
                color=color,
                width=1
            )

    doc.save(output_pdf)

def draw_horizontal_lines(input_pdf, output_pdf, y_positions, color=(1, 0, 0)):
    doc = fitz.open(input_pdf)

    for page in doc:
        height = page.rect.height

        for y in y_positions:

            p1 = fitz.Point(0, y)
            p2 = fitz.Point(height, y)

            page.draw_line(
                p1,
                p2,
                color=color,
                width=1
            )

    doc.save(output_pdf)

def draw_box(input_pdf, output_pdf, bbox, color=(1, 0, 0)):
    doc = fitz.open(input_pdf)
    page = doc[0]

    rect = fitz.Rect(bbox)

    page.draw_rect(
        rect,
        color=color,
        width=0.5
    )

    doc.save(output_pdf)

def draw_boxes(input_pdf, output_pdf, bboxes, color=(1, 0, 0)):
    doc = fitz.open(input_pdf)
    page = doc[0]

    for bbox in bboxes:
        rect = fitz.Rect(bbox)

        page.draw_rect(
            rect,
            color=color,
            width=0.5
        )

    doc.save(output_pdf)

def draw_word_boxes(input_pdf, output_pdf, words):
    doc = fitz.open(input_pdf)
    page = doc[0]

    for w in words:

        rect = fitz.Rect(w["x0"], w["top"], w["x1"], w["bottom"])

        page.draw_rect(
            rect,
            color=(0,0,1),
            width=0.5
        )

    doc.save(output_pdf)
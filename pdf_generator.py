import matplotlib.pyplot as plt
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from io import BytesIO
import os

def create_plot_image(data):
    """Create a matplotlib plot and return it as a BytesIO stream."""
    fig, ax = plt.subplots(figsize=(3, 2))
    ax.plot(data)
    ax.set_title("Sample Chart")
    fig.tight_layout()

    image_stream = BytesIO()
    plt.savefig(image_stream, format='png', bbox_inches='tight')
    plt.close(fig)
    image_stream.seek(0)
    return image_stream

def generate_pdf(plots, comments, output_file):
    c = canvas.Canvas(output_file, pagesize=landscape(A4))
    width, height = landscape(A4)

    positions = [
        (90, 410), (350, 410), (610, 410),   # Top row
        (90, 240), (350, 240), (610, 240),   # Middle row
        (90, 70),  (350, 70),  (610, 70),    # Bottom row
    ]

    for i, (plot_stream, comment) in enumerate(zip(plots, comments)):
        x, y = positions[i]
        c.drawImage(ImageReader(plot_stream), x, y, width=150, height=100)
        c.setFont("Helvetica", 8)
        c.drawString(x, y - 12, comment)

    c.showPage()
    c.save()

def main():
    os.makedirs("output", exist_ok=True)

    comments = [f"Comment for chart {i+1}" for i in range(9)]
    plots = [create_plot_image([i, i+1, i+2]) for i in range(9)]

    generate_pdf(plots, comments, "output/company_report_v2.pdf")

if __name__ == "__main__":
    main()


# import matplotlib.pyplot as plt
# from matplotlib.backends.backend_pdf import PdfPages
# from reportlab.pdfgen import canvas
# from reportlab.lib.pagesizes import A4, landscape
# from reportlab.lib.utils import ImageReader
# import os
#
# def generate_plot(data, filename):
#     # Simple placeholder plot
#     plt.figure(figsize=(3, 2))
#     plt.plot(data)
#     plt.title("Sample Chart")
#     plt.tight_layout()
#     plt.savefig(filename)
#     plt.close()
#
# def generate_pdf(plots, comments, output_file):
#     c = canvas.Canvas(output_file, pagesize=landscape(A4))
#     width, height = landscape(A4)
#
#     # Example coordinates for layout
#     positions = [
#         (90, 410), (350, 410), (610, 410),   # Top row
#         (90, 240), (350, 240), (610, 240),   # Middle row
#         (90, 70),  (350, 70),  (610, 70),    # Bottom row
#     ]
#
#     for i, (plot_path, comment) in enumerate(zip(plots, comments)):
#         x, y = positions[i]
#         c.drawImage(ImageReader(plot_path), x, y, width=150, height=100)
#         c.setFont("Helvetica", 8)
#         c.drawString(x, y - 12, comment)
#
#     c.showPage()
#     c.save()
#
# def main():
#     # Simulated data
#     plots = []
#     comments = [f"Comment for chart {i+1}" for i in range(9)]
#     os.makedirs("plots", exist_ok=True)
#     os.makedirs("output", exist_ok=True)
#
#     for i in range(9):
#         plot_file = f"plots/plot_{i}.png"
#         generate_plot([i, i+1, i+2], plot_file)
#         plots.append(plot_file)
#
#     generate_pdf(plots, comments, "output/company_report.pdf")
#
# if __name__ == "__main__":
#     main()
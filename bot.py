from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
import openpyxl
import csv
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

BOT_TOKEN = "8557795453:AAFZwE9k8JsOxstMpre6My7EuHx12I2OYEo"
EXCEL_FILE = "ict_results.xlsx"
LOG_FILE = "access_log.csv"


# --------- LOGGING ----------
def log_access(user, student_id, status):
    with open(LOG_FILE, "a", newline="") as file:
        writer = csv.writer(file)
        writer.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            user,
            student_id,
            status
        ])


# --------- FETCH RESULT ----------
def get_result(student_id):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    sheet = wb["results"]

    for row in sheet.iter_rows(min_row=2, values_only=True):
        sid, full_name, grade_section, semester, year, ca, final, total = row

        if sid and sid.strip().lower() == student_id.strip().lower():
            return {
                "student_id": sid,
                "name": full_name,
                "grade_section": grade_section,
                "semester": semester,
                "year": year,
                "ca": ca,
                "final": final,
                "total": total
            }
    return None


# --------- PDF CREATION ----------
def create_pdf(result):
    filename = f"{result['student_id']}_report.pdf"
    c = canvas.Canvas(filename, pagesize=A4)

    c.setFont("Helvetica", 14)
    c.drawString(100, 800, "ICT Report Card")

    c.setFont("Helvetica", 12)
    c.drawString(100, 760, f"Student ID: {result['student_id']}")
    c.drawString(100, 740, f"Full Name: {result['name']}")
    c.drawString(100, 720, f"Grade & Section: {result['grade_section']}")
    c.drawString(100, 700, f"Semester: {result['semester']}")
    c.drawString(100, 680, f"Year: {result['year']}")

    c.drawString(100, 640, f"Continuous Assessment: {result['ca']}")
    c.drawString(100, 620, f"Final Exam: {result['final']}")
    c.drawString(100, 600, f"Total: {result['total']}")

    c.save()
    return filename


# --------- TELEGRAM COMMANDS ----------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Welcome 👋\n\n"
        "Dilla Don Bosco Secondary School\n\n"
        "Send your Student ID to view your Information Technology result.\n"
        "Please keep your Student ID private."

        

    )


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    student_id = update.message.text.strip()
    user = update.message.from_user.username

    result = get_result(student_id)

    if result:
        log_access(user, student_id, "FOUND")

        msg = f"""
📘 Information Technology Exam Result


Full Name: {result['name']}
Grade & Section: {result['grade_section']}
Semester: {result['semester']}
Year: {result['year']}

Continuous Assessment: {result['ca']}
Final Exam: {result['final']}
Total: {result['total']}
"""

        context.user_data["result"] = result
        await update.message.reply_text(msg)

    else:
        log_access(user, student_id, "NOT FOUND")
        await update.message.reply_text("❗ Student ID not found.")


async def send_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "result" not in context.user_data:
        await update.message.reply_text("❗ First send your Student ID.")
        return

    filename = create_pdf(context.user_data["result"])
    await update.message.reply_document(open(filename, "rb"))


# --------- BOT SETUP ----------
app = ApplicationBuilder().token(BOT_TOKEN).build()

app.add_handler(CommandHandler("start", start))
app.add_handler(CommandHandler("pdf", send_pdf))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

print("Bot running...")
app.run_polling()

from docxtpl import DocxTemplate, InlineImage

with open("work_sheet.csv","r",encoding='utf-8') as csvf:
    op = csvf.readlines()
print(type(op))

for i in op:

    todayStr = i.split(",")[0]
    recipientName = i.split(",")[1]
    evntDtStr = i.split(",")[2]
    venueStr = i.split(",")[3]
    senderName = i.split(",")[4]
    print(f"{todayStr}{recipientName}{evntDtStr}{venueStr}{senderName}")

    doc = DocxTemplate("inviteTemp.docx")
    context = {
        "todayStr":todayStr,
        "recipientName":recipientName,
        "evntDtStr":evntDtStr,
        "venueStr":venueStr,
        "senderName":senderName,
        "bannering": InlineImage(doc, "party.jpg")
    }
    doc.render(context)
    doc.save(f"invites/invitation_{recipientName}.docx")


def delete_unwanted_spaces(text):

    if text[0] == " ":
        text = text[1:]
    if text[-1] == " ":
        text = text[:-1]

    text = text.replace("   ", " ")
    text = text.replace("  ", " ")

    # تصحيح الأسماء التعبيدية
    text = text.replace("عبدالله", "عبد الله")
    text = text.replace("عبدالعزيز", "عبد العزيز")
    text = text.replace("عبدالرحمن", "عبد الرحمن")
    text = text.replace("عبدالخالق", "عبد الخالق")
    text = text.replace("عبدالقادر", "عبد القادر")
    text = text.replace("عبدالحق", "عبد الحق")
    text = text.replace("عبدالواحد", "عبد الواحد")
    text = text.replace("عبدالكريم", "عبد الكريم")
    text = text.replace("عبدالغني", "عبد الغني")
    text = text.replace("عبدالاله", "عبد الإله")

    # تعديل التاء المربوطة
    text = text.replace("منيره", "منيرة")
    text = text.replace("فاطمه", "فاطمة")
    text = text.replace("حصه", "حصة")
    text = text.replace("ساره", "سارة")
    text = text.replace("صفيه", "صفية")
    text = text.replace("رقيه", "رقية")
    text = text.replace("هيله", "هيلة")
    text = text.replace("ميمونه", "ميمونة")
    text = text.replace("لولوه", "لولوة")
    text = text.replace("مزنه", "مزنة")
    text = text.replace("بدريه", "بدرية")
    text = text.replace("شريفه", "شريفة")
    text = text.replace("اميه", "أمية")
    text = text.replace("أميه", "أمية")

    # همزة القطع والوصل
    text = text.replace("احمد", "أحمد")
    text = text.replace("ابراهيم", "إبراهيم")
    text = text.replace("امل", "أمل")
    text = text.replace("الزأمل", "الزامل")
    text = text.replace("اريج", "أريج")

    return text

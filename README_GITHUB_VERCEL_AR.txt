Hapag-Lloyd Egypt Manifest Checker — GitHub + Vercel

الرفع:
1) فك ضغط ZIP.
2) ارفع محتويات المجلد إلى Repository جديد على GitHub (يجب أن يكون index.html في الجذر).
3) من Vercel اختر Add New Project ثم Import Git Repository.
4) Framework Preset: Other، ثم Deploy.

ملاحظات مهمة وصريحة:
- هذه نسخة تجريبية Serverless تستعمل Python وFlask داخل api/index.py.
- Vercel لها حد لحجم الطلب (عادة نحو 4.5MB) وحد زمني للتنفيذ. ملفات PDF الكبيرة (مثل 100MB أو مئات الصفحات) قد تفشل، حتى لو كان الكود صحيحًا.
- للمانيفستات الكبيرة والاعتمادية الإنتاجية، يلزم رفع PDF إلى تخزين خارجي (مثلاً S3/R2) ومعالجته بواسطة worker طويل التشغيل (Render/Railway/VPS)، ثم إرجاع Excel. لا يمكن ضمان معالجة 480+ صفحة على Vercel Function وحدها.
- بعض إعدادات حساب Vercel قد لا تدعم Python Functions أو مدة 60 ثانية؛ عندها نحتاج استضافة backend منفصلة.

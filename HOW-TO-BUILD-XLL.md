# إزاي تجيب XLL جديد (زي PlanningToolkit-v0_6_2-x64.xll) للنسخة v0.7.0

الملف اللي رفعته (`PlanningToolkit-v0_6_2-x64.xll`) هو ناتج بناء تلقائي من GitHub Actions
(`windows-build.yml`) — نفس الآلية دي هي اللي هتجيبلك نسخة v0.7.0 بالـ theme الجديد.

أنا مش قادر أبني ملف .xll بنفسي هنا لأن ده ملف Windows (PE32+ / Excel-DNA add-in) ومحتاج
Excel COM وWindows Forms عشان يتبني — والبيئة اللي بشتغل فيها Linux من غير Windows أو .NET SDK.
لكن جهزت الكود بحيث لما تشغل نفس الـ workflow هيطلعلك XLL جاهز بنفس الطريقة بالظبط.

## الخطوات

1. **استبدل الملفات دي في الريبو** (نفس الأسماء، في الـ root):
   كل ملفات `.cs` و`.csproj` الموجودة في هذا المجلد، بالإضافة لـ `ReportTheme.cs`
   و`XerReportExport.cs` (دول جداد).

2. **احذف أو أرشف باقي ملفات الـ Patch القديمة** من الـ root:
   `PlanningToolkit-Phase2-Patch.zip` ... `PlanningToolkit-Phase6-Patch.zip`
   (لو سبتهم موجودين، ولا حاجة هتحصل — الـ workflow الجديد مبقاش بيفكهم، لكن الأفضل تشيلهم
   أو تحطهم في مجلد `archive/` عشان الريبو يفضل نضيف).

3. **استبدل** `.github/workflows/windows-build.yml` بالنسخة الجديدة الموجودة هنا
   (فيها إضافة `ReportTheme.cs` و`XerReportExport.cs` لقائمة الملفات اللي بتتنقل لـ
   `src/PlanningToolkit.Excel/`، وترقيم الإصدار v0.7.0).

4. **Commit و Push** على branch `main`.

5. الـ workflow هيشتغل تلقائي (أو شغّله يدوي من تبويب Actions → Build Windows XLL →
   Run workflow). لما يخلص، هتلاقي artifact اسمه
   `PlanningToolkit-Windows-x64-v0.7.0` — نزّله، جواه `PlanningToolkit-v0.7.0-x64.xll`
   بالظبط زي اللي رفعته بس بالتحديثات الجديدة.

## لو حصل خطأ في الـ build

ابعتلي رسالة الخطأ من الـ Actions log وهحلها فورًا — الكود اتراجع سطر بسطر ومفيش أخطاء
واضحة، لكن التأكيد الحقيقي بيحصل وقت الـ build على Windows.

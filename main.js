 function processExcel() {
            const fileInput = document.getElementById('excelFile');
            const oldDateInput = document.getElementById('oldDate').value.trim();
            const newDateInput = document.getElementById('newDate').value.trim();
            const messageDiv = document.getElementById('message');

            if (!fileInput.files.length) {
                alert("الرجاء اختيار ملف إكسل أولاً."); return;
            }
            if (!oldDateInput || !newDateInput) {
                alert("الرجاء إدخال التاريخ القديم والجديد."); return;
            }

            messageDiv.style.color = "#007bff";
            messageDiv.innerText = "جاري معالجة الملف...";

            const file = fileInput.files[0];
            const reader = new FileReader();

            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {type: 'array'});
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    let jsonData = XLSX.utils.sheet_to_json(worksheet);

                    // البحث والاستبدال في عمود name
                    jsonData = jsonData.map(row => {
                        if (row['name']) {
                            const regex = new RegExp(oldDateInput, 'g');
                            row['name'] = row['name'].replace(regex, newDateInput);
                        }
                        return row;
                    });

                    const newWorksheet = XLSX.utils.json_to_sheet(jsonData);
                    workbook.Sheets[firstSheetName] = newWorksheet;
                    
                    const newFileName = `new_file_${newDateInput}.xlsx`;
                    XLSX.writeFile(workbook, newFileName);

                    messageDiv.style.color = "green";
                    messageDiv.innerText = "تم التحديث والتحميل بنجاح!";
                } catch (error) {
                    console.error(error);
                    messageDiv.style.color = "red";
                    messageDiv.innerText = "حدث خطأ أثناء المعالجة.";
                }
            };
            reader.readAsArrayBuffer(file);
        }
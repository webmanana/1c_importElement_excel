<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Обработка данных из Excel</title>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f7f9fc;
        }
        h1 {
            color: #333;
        }
        #status {
            margin: 10px 0;
            font-size: 16px;
            color: #555;
        }
        #start {
            padding: 10px 20px;
            background-color: #007bff;
            color: #fff;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 16px;
        }
        #start:disabled {
            background-color: #6c757d;
            cursor: not-allowed;
        }
        .error {
            color: red;
            font-size: 14px;
            margin-top: 10px;
        }
    </style>
</head>
<body>
    <h1>Обработка данных из Excel</h1>
    <p id="status">Нажмите кнопку ниже для начала обработки данных.</p>
    <button id="start">Начать обработку</button>

    <script>
        let position = 0;

        function processData() {
            $.ajax({
                url: '/local/scripts/process_excel_step.php',
                method: 'POST',
                data: { position: position },
                dataType: 'json',
                timeout: 60000, // Таймаут 60 секунд
                success: function(response) {
                    if (response.status === 'done') {
                        $('#status').text('Обработка данных успешно завершена!');
                        $('#start').text('Завершено').attr('disabled', true);
                    } else if (response.status === 'in_progress') {
                        position = response.next_position;
                        $('#status').text(`Обработано ${position} строк...`);
                        processData(); // Запуск следующей итерации
                    }
                },
                error: function(jqXHR, textStatus, errorThrown) {
                    const errorMessage = `
                        Произошла ошибка при обработке данных:
                        <br>Код ошибки: ${jqXHR.status} (${jqXHR.statusText})
                        <br>Ошибка: ${errorThrown}
                        <br>Ответ сервера: ${jqXHR.responseText || 'Нет ответа'}
                    `;
                    $('#status').html(`<div class="error">${errorMessage}</div>`); // Вывод ошибки
                    console.error('Детали ошибки:', {
                        status: jqXHR.status,
                        statusText: jqXHR.statusText,
                        errorThrown: errorThrown,
                        responseText: jqXHR.responseText
                    });
                    $('#start').text('Повторить обработку').attr('disabled', false);
                }
            });
        }

        $(document).ready(function() {
            $('#start').on('click', function() {
                $(this).attr('disabled', true).text('Обработка в процессе...');
                processData();
            });
        });
    </script>
</body>
</html>

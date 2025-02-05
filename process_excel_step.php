<?php
require($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/prolog_before.php");
require $_SERVER['DOCUMENT_ROOT'] . '/local/scripts/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use Bitrix\Main\Loader;
use Bitrix\Iblock\ElementTable;

Loader::includeModule("iblock");

$iblockId = 10; // Замените на актуальный ID вашего инфоблока

// Проверяем, существует ли инфоблок
$iblockExists = \CIBlock::GetList([], ["ID" => $iblockId])->Fetch();
if (!$iblockExists) {
    echo json_encode(["status" => "error", "message" => "Инфоблок с ID $iblockId не найден."]);
    exit;
}

// Путь к файлу обработки
$filePath = $_SERVER['DOCUMENT_ROOT'] . "/local/scripts/act.xlsx";
$batchSize = 100;
$currentPosition = isset($_POST['position']) ? (int)$_POST['position'] : 0;

// Загружаем данные из Excel
$spreadsheet = IOFactory::load($filePath);
$sheet = $spreadsheet->getActiveSheet();

$data = [];
foreach ($sheet->getRowIterator($currentPosition + 2, $currentPosition + $batchSize + 1) as $row) {
    $rowData = [];
    foreach ($row->getCellIterator() as $cell) {
        $rowData[] = $cell->getValue();
    }
    $data[] = $rowData;
}

if (empty($data)) {
    echo json_encode(["status" => "done"]);
    exit;
}

// Обрабатываем элементы инфоблока
foreach ($data as $row) {
    $identifier = $row[1] ?? null;
    if (!$identifier) continue; // Пропускаем пустые строки

    $existingElement = ElementTable::getList([
        'select' => ['ID'],
        'filter' => ['IBLOCK_ID' => $iblockId, 'NAME' => $identifier],
        'limit' => 1
    ])->fetch();

    if ($existingElement) {
        continue; // Если элемент уже существует, пропускаем его
    }

    // Добавляем новый элемент
    $el = new \CIBlockElement;
    $el->Add([
        "IBLOCK_ID" => $iblockId,
        "NAME" => $identifier,
        "ACTIVE" => "Y",
        "ACTIVE_FROM" => $row[2] ?? "",
        "PROPERTY_VALUES" => [
            "START_DATE" => $row[3] ?? "",
            "END_DATE" => $row[4] ?? "",
            "VARIANT" => $row[6] ?? "",
            "REWARD" => isset($row[7]) ? str_replace(',', '.', $row[7]) : "", // Заменяем запятую на точку
            "CODE" => $row[8] ?? ""
        ]
    ]);
}

echo json_encode(["status" => "in_progress", "next_position" => $currentPosition + $batchSize]);

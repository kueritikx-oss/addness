<?php
// ============================================
// Sales Timer - Google Sheets Proxy
// スプレッドシートのデータをJSON APIとして提供
// ============================================
error_reporting(0);
ini_set('display_errors', 0);
header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');
header('Cache-Control: no-cache');

// --- 設定 ---
$CONFIG_FILE = __DIR__ . '/.sheets-token.json';
$CLIENT_ID = '202264815644.apps.googleusercontent.com';
$CLIENT_SECRET = 'X4Z3ca8xfWDb1Voo-F9a7ZxJ';
$SPREADSHEET_ID = '1J0cKmd522tWeMo2NOl9QNHKCFaKe3UoRR6A002yM0Xo';

// --- トークン管理 ---
function loadToken() {
    global $CONFIG_FILE;
    if (!file_exists($CONFIG_FILE)) {
        return null;
    }
    return json_decode(file_get_contents($CONFIG_FILE), true);
}

function saveToken($token) {
    global $CONFIG_FILE;
    file_put_contents($CONFIG_FILE, json_encode($token, JSON_PRETTY_PRINT));
}

function refreshAccessToken() {
    global $CLIENT_ID, $CLIENT_SECRET;
    $token = loadToken();
    if (!$token || empty($token['refresh_token'])) {
        return null;
    }

    $ch = curl_init('https://oauth2.googleapis.com/token');
    curl_setopt_array($ch, [
        CURLOPT_POST => true,
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_POSTFIELDS => http_build_query([
            'grant_type' => 'refresh_token',
            'client_id' => $CLIENT_ID,
            'client_secret' => $CLIENT_SECRET,
            'refresh_token' => $token['refresh_token']
        ])
    ]);
    $res = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    if ($httpCode !== 200) {
        return null;
    }

    $data = json_decode($res, true);
    if (!empty($data['access_token'])) {
        $token['access_token'] = $data['access_token'];
        $token['expiry'] = time() + ($data['expires_in'] ?? 3600);
        saveToken($token);
        return $token['access_token'];
    }
    return null;
}

function getAccessToken() {
    $token = loadToken();
    if (!$token) return null;

    // トークンが有効期限内ならそのまま使う
    if (!empty($token['access_token']) && !empty($token['expiry']) && $token['expiry'] > time() + 60) {
        return $token['access_token'];
    }

    // リフレッシュ
    return refreshAccessToken();
}

// --- gviz でシートデータを取得 ---
function fetchSheet($spreadsheetId, $sheetName, $accessToken, $query = null) {
    $url = "https://docs.google.com/spreadsheets/d/{$spreadsheetId}/gviz/tq?tqx=out:csv";
    $url .= "&sheet=" . urlencode($sheetName);
    if ($query) {
        $url .= "&tq=" . urlencode($query);
    }

    $ch = curl_init($url);
    curl_setopt_array($ch, [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_FOLLOWLOCATION => true,
        CURLOPT_HTTPHEADER => [
            "Authorization: Bearer {$accessToken}"
        ]
    ]);
    $res = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    if ($httpCode !== 200 || strpos($res, '<!DOCTYPE') !== false) {
        return null;
    }
    return $res;
}

// --- CSV パース ---
function parseCSV($csv) {
    $rows = [];
    $stream = fopen('php://temp', 'r+');
    fwrite($stream, $csv);
    rewind($stream);
    while (($row = fgetcsv($stream)) !== false) {
        $rows[] = $row;
    }
    fclose($stream);
    return $rows;
}

// --- 列文字 → インデックス ---
function colToIndex($col) {
    $col = strtoupper(trim($col));
    $idx = 0;
    for ($i = 0; $i < strlen($col); $i++) {
        $idx = $idx * 26 + (ord($col[$i]) - 64);
    }
    return $idx - 1;
}

// ============================================
// メイン処理
// ============================================

// トークンが未設定の場合
$token = loadToken();
if (!$token) {
    echo json_encode(['error' => 'トークン未設定。セットアップが必要です。', 'setup_required' => true]);
    exit;
}

$accessToken = getAccessToken();
if (!$accessToken) {
    echo json_encode(['error' => 'アクセストークンの取得に失敗しました。再認証が必要です。']);
    exit;
}

// クエリパラメータ
$from = $_GET['from'] ?? null;
$to = $_GET['to'] ?? null;

try {
    // 1. 調整シートから列設定を取得
    $configCSV = fetchSheet($SPREADSHEET_ID, '調整', $accessToken, 'select E');
    $configRows = $configCSV ? parseCSV($configCSV) : [];

    // E6, E12, E15 (CSVでは0ベースでヘッダー含む: 行5, 11, 14)
    $dateCol = isset($configRows[5][0]) ? trim($configRows[5][0]) : 'C';
    $amountCol = isset($configRows[11][0]) ? trim($configRows[11][0]) : 'F';
    $statusCol = isset($configRows[14][0]) ? trim($configRows[14][0]) : 'G';

    // 2. 入金管理リストのデータ取得
    $dataCSV = fetchSheet($SPREADSHEET_ID, '入金管理リスト', $accessToken);
    if (!$dataCSV) {
        throw new Exception('入金管理リストのデータ取得に失敗しました');
    }

    $rows = parseCSV($dataCSV);
    if (count($rows) < 2) {
        throw new Exception('データが空です');
    }

    $dateIdx = colToIndex($dateCol);
    $amountIdx = colToIndex($amountCol);
    $statusIdx = colToIndex($statusCol);

    $fromDate = $from ? strtotime($from) : null;
    $toDate = $to ? strtotime($to . ' 23:59:59') : null;

    $records = [];
    $total = 0;

    // ヘッダー行をスキップ (i=1から)
    for ($i = 1; $i < count($rows); $i++) {
        $row = $rows[$i];
        $status = trim($row[$statusIdx] ?? '');
        $dateStr = $row[$dateIdx] ?? '';

        if ($status === '') continue;

        // 日付フィルタ
        if ($dateStr && ($fromDate || $toDate)) {
            $rowDate = strtotime($dateStr);
            if ($rowDate !== false) {
                if ($fromDate && $rowDate < $fromDate) continue;
                if ($toDate && $rowDate > $toDate) continue;
            }
        }

        if ($status === '成約' || strpos($status, '契約済み') !== false) {
            $rawAmount = preg_replace('/[,¥￥\s]/', '', $row[$amountIdx] ?? '0');
            $amount = floatval($rawAmount);

            if ($amount > 0) {
                $records[] = [
                    'date' => $dateStr,
                    'amount' => $amount,
                    'status' => $status
                ];
                $total += $amount;
            }
        }
    }

    // 3. 個人成績シートから成約率を取得
    $summaryData = [
        'closingRate' => null,
        'reservations' => null,
        'conducted' => null,
        'closed' => null
    ];
    $summaryCSV = fetchSheet($SPREADSHEET_ID, '個人成績', $accessToken);
    if ($summaryCSV) {
        $summaryRows = parseCSV($summaryCSV);
        // 「契約締結日ベース」の月間サマリー行を探す
        for ($i = 0; $i < count($summaryRows) - 2; $i++) {
            $cell = trim($summaryRows[$i][0] ?? '');
            // 「契約締結日ベース」のヘッダーを見つけたらその2行下がデータ
            if (strpos($cell, '契約') !== false && strpos($cell, '締結日') !== false && strpos($cell, 'ベース') !== false) {
                $dataRow = $summaryRows[$i + 2] ?? null;
                if ($dataRow) {
                    $summaryData['reservations'] = intval($dataRow[0] ?? 0);
                    $summaryData['conducted'] = intval($dataRow[3] ?? 0);
                    $summaryData['closed'] = intval($dataRow[9] ?? 0);
                    // K列(index 10)に成約率がある
                    $rateStr = trim($dataRow[10] ?? '');
                    if ($rateStr !== '' && $rateStr !== '-') {
                        $summaryData['closingRate'] = floatval(str_replace('%', '', $rateStr));
                    }
                }
                break;
            }
        }
    }

    echo json_encode([
        'total' => $total,
        'records' => $records,
        'count' => count($records),
        'summary' => $summaryData,
        'columns' => ['date' => $dateCol, 'amount' => $amountCol, 'status' => $statusCol],
        'timestamp' => date('c')
    ], JSON_UNESCAPED_UNICODE);

} catch (Exception $e) {
    echo json_encode(['error' => $e->getMessage()], JSON_UNESCAPED_UNICODE);
}

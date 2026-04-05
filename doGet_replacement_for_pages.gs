function doGet(e) {
  if (e && e.parameter && e.parameter.test === '1') {
    return HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>テスト: クリエイター取得</title>
</head>
<body>
  <h1>テスト: クリエイター取得</h1>
  <div id="status">読み込み中...</div>
  <pre id="detail"></pre>
  <script>
    function setResult(status, detail) {
      document.getElementById('status').textContent = status;
      document.getElementById('detail').textContent = detail || '';
    }
    try {
      google.script.run
        .withSuccessHandler(function(creators) {
          var count = Array.isArray(creators) ? creators.length : 0;
          setResult('取得成功: ' + count + '件', JSON.stringify(creators.slice(0, 2), null, 2));
        })
        .withFailureHandler(function(err) {
          setResult('取得失敗', JSON.stringify(err, null, 2));
        })
        .getCreators();
    } catch (err) {
      setResult('実行時エラー', String(err));
    }
  </script>
</body>
</html>
`)
      .setTitle('テスト: クリエイター取得')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (e && e.parameter && e.parameter.mode === 'json') {
    const data = getCreators();

    if (e.parameter.callback) {
      const callback = String(e.parameter.callback).replace(/[^\w.$]/g, '');
      return ContentService
        .createTextOutput(callback + '(' + JSON.stringify(data) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const template = HtmlService.createTemplateFromFile('index');
  template.apiKey = MAPS_API_KEY;
  return template.evaluate()
    .setTitle('千葉ニュータウン クリエイターズマップ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

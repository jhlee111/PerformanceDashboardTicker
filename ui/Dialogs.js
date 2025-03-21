/**
 * Performance Dashboard Ticker - Dialog UI Module
 * 
 * This module contains functions for creating and managing UI dialogs.
 */

/**
 * Show a dialog for setting the reference date
 */
function showReferenceDateDialog() {
  try {
    Logger.log('참조일 설정 대화상자를 표시합니다...');
    
    // Get current reference date
    const currentRefDate = getReferenceDate();
    const formattedDate = Utilities.formatDate(currentRefDate, "GMT+9", "yyyy-MM-dd");
    
    // Create HTML for the dialog
    const htmlOutput = HtmlService
      .createHtmlOutput(
        `<style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
          }
          .form-group {
            margin-bottom: 15px;
          }
          label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
          }
          input[type="date"] {
            width: 100%;
            padding: 8px;
            box-sizing: border-box;
            border: 1px solid #ccc;
            border-radius: 4px;
          }
          .buttons {
            display: flex;
            justify-content: flex-end;
            margin-top: 20px;
          }
          button {
            padding: 8px 16px;
            margin-left: 10px;
            cursor: pointer;
            background-color: #4285f4;
            color: white;
            border: none;
            border-radius: 4px;
          }
          button.cancel {
            background-color: #f1f1f1;
            color: #333;
          }
          .note {
            font-size: 12px;
            color: #666;
            margin-top: 5px;
          }
        </style>
        <div>
          <h2>참조일 설정</h2>
          <p>대시보드 계산의 기준이 되는 참조일을 설정합니다.</p>
          
          <div class="form-group">
            <label for="refDate">참조일:</label>
            <input type="date" id="refDate" value="${formattedDate}">
            <div class="note">참조일을 기준으로 주간, 월간, YTD 수익률이 계산됩니다.</div>
          </div>
          
          <div class="buttons">
            <button class="cancel" onclick="google.script.host.close()">취소</button>
            <button onclick="setDate()">저장</button>
          </div>
        </div>
        
        <script>
          function setDate() {
            const dateValue = document.getElementById('refDate').value;
            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .setReferenceDate(dateValue);
          }
        </script>`)
      .setWidth(400)
      .setHeight(300)
      .setTitle('참조일 설정');
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '참조일 설정');
  } catch (error) {
    Logger.log(`참조일 설정 대화상자 표시 오류: ${error.message}`);
    SpreadsheetApp.getUi().alert('오류', `참조일 설정 대화상자를 표시할 수 없습니다: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Show the ticker manager dialog
 */
function showTickerManager() {
  try {
    Logger.log('티커 관리 대화상자를 표시합니다...');
    
    // Open the tickers sheet directly
    openTickerSheet();
    
    // Show a simple help message
    SpreadsheetApp.getUi().alert(
      '티커 관리',
      '티커 시트에서 티커를 관리할 수 있습니다.\n\n' +
      '각 행에 티커 정보를 입력하거나 수정하고 대시보드를 업데이트하면 변경 사항이 반영됩니다.\n\n' +
      '컬럼 형식:\n' +
      '- 이름: 표시할 티커 이름\n' +
      '- 티커: 심볼 코드 (예: AAPL, 005930.KS)\n' +
      '- 소스: 데이터 소스 (google, yahoo, naver)',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    Logger.log(`티커 관리 대화상자 표시 오류: ${error.message}`);
    SpreadsheetApp.getUi().alert('오류', `티커 관리 대화상자를 표시할 수 없습니다: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Show the help dialog
 */
function showHelp() {
  try {
    Logger.log('도움말을 표시합니다...');
    
    const htmlOutput = HtmlService
      .createHtmlOutput(
        `<style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            color: #333;
            line-height: 1.5;
          }
          h1 {
            color: #1a73e8;
            font-size: 24px;
            margin-top: 0;
            border-bottom: 1px solid #eee;
            padding-bottom: 10px;
          }
          h2 {
            color: #188038;
            font-size: 18px;
            margin-top: 20px;
            margin-bottom: 10px;
          }
          p {
            margin-bottom: 15px;
          }
          ul {
            padding-left: 25px;
          }
          li {
            margin-bottom: 8px;
          }
          .note {
            background-color: #f8f9fa;
            padding: 10px;
            border-left: 4px solid #1a73e8;
            margin: 15px 0;
          }
          .warning {
            background-color: #fef7e0;
            padding: 10px;
            border-left: 4px solid #f9ab00;
            margin: 15px 0;
          }
          code {
            font-family: monospace;
            background-color: #f1f3f4;
            padding: 2px 5px;
            border-radius: 3px;
          }
        </style>
        <h1>성과 대시보드 도움말</h1>
        
        <h2>기본 사용법</h2>
        <p>성과 대시보드는 다양한 금융 지수와 종목의 성과를 추적하고 표시합니다. 다음과 같은 기능을 제공합니다:</p>
        <ul>
          <li><strong>대시보드 업데이트:</strong> 최신 가격 데이터로 대시보드를 업데이트합니다.</li>
          <li><strong>참조일 설정:</strong> 수익률 계산의 기준이 될 참조일을 설정합니다.</li>
          <li><strong>티커 관리:</strong> 대시보드에 표시할 금융 지수와 종목을 관리합니다.</li>
        </ul>
        
        <h2>티커 관리</h2>
        <p>티커 관리 메뉴를 클릭하면 티커 시트로 이동합니다. 이 시트에서 다음 정보를 입력하여 티커를 관리할 수 있습니다:</p>
        <ul>
          <li><strong>이름:</strong> 대시보드에 표시될 종목 이름</li>
          <li><strong>티커:</strong> 데이터 소스에서 사용하는 심볼 코드 (예: AAPL, 005930.KS)</li>
          <li><strong>소스:</strong> 데이터를 가져올 소스 (google, yahoo, naver 중 하나)</li>
        </ul>
        
        <div class="note">
          <p><strong>참고:</strong> 티커 정보를 변경한 후에는 대시보드 업데이트를 실행하여 변경 사항을 적용해야 합니다.</p>
        </div>
        
        <h2>참조일 설정</h2>
        <p>참조일은 수익률 계산의 기준이 되는 날짜입니다. 다음과 같은 수익률이 계산됩니다:</p>
        <ul>
          <li><strong>주간 수익률:</strong> 참조일로부터 7일 전 대비 수익률</li>
          <li><strong>월간 수익률:</strong> 참조일로부터 30일 전 대비 수익률</li>
          <li><strong>YTD 수익률:</strong> 참조일이 속한 연도의 1월 1일 대비 수익률</li>
          <li><strong>고점 대비 수익률:</strong> 52주 고점 대비 수익률</li>
        </ul>
        
        <h2>지원되는 데이터 소스</h2>
        <p>다음 데이터 소스를 지원합니다:</p>
        <ul>
          <li><strong>Google Finance:</strong> Google 금융 데이터</li>
          <li><strong>Yahoo Finance:</strong> Yahoo 금융 데이터</li>
          <li><strong>Naver Finance:</strong> 네이버 금융 데이터 (주로 한국 주식 시장용)</li>
        </ul>
        
        <div class="warning">
          <p><strong>주의:</strong> 대시보드 업데이트 중에는 다른 업데이트를 실행할 수 없습니다. 이미 업데이트가 진행 중인 경우 완료될 때까지 기다려야 합니다.</p>
          <p>만약 업데이트가 장시간 진행되거나 완료되지 않는 경우, [📊 대시보드 > 🛠️ 진단 도구 > ⚠️ 스크립트 강제 초기화] 메뉴를 사용하여 시스템을 초기화할 수 있습니다.</p>
        </div>
        `)
      .setWidth(600)
      .setHeight(500)
      .setTitle('성과 대시보드 도움말');
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '도움말');
  } catch (error) {
    Logger.log(`도움말 표시 오류: ${error.message}`);
    SpreadsheetApp.getUi().alert('오류', `도움말을 표시할 수 없습니다: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Show the symbol diagnostics dialog
 */
function showSymbolDiagnosticsDialog() {
  try {
    Logger.log('심볼 진단 대화상자를 표시합니다...');
    
    const htmlOutput = HtmlService
      .createHtmlOutput(
        `<style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
          }
          .form-group {
            margin-bottom: 15px;
          }
          label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
          }
          input, select {
            width: 100%;
            padding: 8px;
            box-sizing: border-box;
            border: 1px solid #ccc;
            border-radius: 4px;
          }
          .buttons {
            display: flex;
            justify-content: flex-end;
            margin-top: 20px;
          }
          button {
            padding: 8px 16px;
            margin-left: 10px;
            cursor: pointer;
            background-color: #4285f4;
            color: white;
            border: none;
            border-radius: 4px;
          }
          button.cancel {
            background-color: #f1f1f1;
            color: #333;
          }
          .note {
            font-size: 12px;
            color: #666;
            margin-top: 5px;
          }
        </style>
        <div>
          <h2>심볼 진단</h2>
          <p>특정 심볼에 대한 진단을 실행합니다.</p>
          
          <div class="form-group">
            <label for="symbol">심볼:</label>
            <input type="text" id="symbol" placeholder="AAPL, KRX:005930, ^KS11">
            <div class="note">예: AAPL, KRX:005930, ^KS11</div>
          </div>
          
          <div class="form-group">
            <label for="source">데이터 소스:</label>
            <select id="source">
              <option value="google">Google Finance</option>
              <option value="yahoo">Yahoo Finance</option>
              <option value="naver">Naver Finance</option>
            </select>
          </div>
          
          <div class="buttons">
            <button class="cancel" onclick="google.script.host.close()">취소</button>
            <button onclick="runDiagnostic()">진단 실행</button>
          </div>
        </div>
        
        <script>
          function runDiagnostic() {
            const symbol = document.getElementById('symbol').value;
            const source = document.getElementById('source').value;
            
            if (!symbol) {
              alert('심볼을 입력해주세요.');
              return;
            }
            
            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .runSymbolDiagnostics(symbol, source);
          }
        </script>`)
      .setWidth(400)
      .setHeight(350)
      .setTitle('심볼 진단');
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '심볼 진단');
  } catch (error) {
    Logger.log(`심볼 진단 대화상자 표시 오류: ${error.message}`);
    SpreadsheetApp.getUi().alert('오류', `심볼 진단 대화상자를 표시할 수 없습니다: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
} 
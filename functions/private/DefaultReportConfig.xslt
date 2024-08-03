<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
   <xsl:param name="includeButtons" select="'true'"/>
  <!-- Template to exclude comments -->
  <xsl:template match="comment()"/>
  <xsl:template match="/">
    <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #1e1e1e;
            color: #fff;
          }
          .container {
            width: 90%;
            margin: auto;
            overflow: hidden;
          }
          header {
            background: #000000;
            color: white;
            padding-top: 30px;
            min-height: 70px;
            border-bottom: #e8491d 3px solid;
          }
          header h1 {
            padding: 5px 0;
            text-align: center;
          }
          .summary-box {
            display: flex;
            justify-content: space-around;
            padding: 20px;
            background-color: #2c2c2c;
            margin-bottom: 20px;
          }
          .summary-item {
            text-align: center;
            flex: 1;
            margin: 10px;
          }
          .summary-item h2 {
            margin: 0;
            font-size: 2em;
          }
          .summary-item p {
            margin: 5px 0 0 0;
            font-size: 1.2em;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            background-color: #333;
            color: #fff;
          }
          th, td {
            padding: 15px;
            text-align: left;
            border-bottom: 1px solid #444;
          }
          .success {
            background-color: #4CAF50;
            color: white;
          }
          .failure {
            background-color: #f44336;
            color: white;
          }
          .not-tested {
            background-color: #777;
            color: white;
          }
          .filter-buttons {
            text-align: center;
            margin-bottom: 20px;
          }
          .filter-buttons button {
            padding: 10px 20px;
            margin: 0 5px;
            background-color: #444;
            color: #fff;
            border: none;
            cursor: pointer;
          }
          .filter-buttons button:hover {
            background-color: #666;
          }
          .nowrap {
            white-space: nowrap;
          }
        </style>
        <script>
          function filterTests(filter) {
            var rows = document.querySelectorAll('table tr.test-row');
            rows.forEach(function(row) {
              if (filter === 'all' || row.classList.contains(filter)) {
                row.style.display = '';
              } else {
                row.style.display = 'none';
              }
              // Add nowrap class to failed messages when filter is 'Passed'
              var failedMessageCell = row.querySelector('.failure-message');
              if (failedMessageCell) {
                if (filter === 'Success') {
                  failedMessageCell.classList.add('nowrap');
                } else {
                  failedMessageCell.classList.remove('nowrap');
                }
              }
            });
          }
        </script>
      </head>
      <body>
        <header>
          <div class="container">
            <h1>365AutomatedCheck Results</h1>
          </div>
        </header>
        <div class="container">
          <div class="summary-box">
            <div class="summary-item">
              <h2><xsl:value-of select="count(//test-case)"/></h2>
              <p>Total tests</p>
            </div>
            <div class="summary-item">
              <h2><xsl:value-of select="count(//test-case[@result='Success'])"/></h2>
              <p>Passed</p>
            </div>
            <div class="summary-item">
              <h2><xsl:value-of select="count(//test-case[@result='Failure'])"/></h2>
              <p>Failed</p>
            </div>
            <div class="summary-item">
              <h2><xsl:value-of select="count(//test-case[not(@result)])"/></h2>
              <p>Not tested</p>
            </div>
          </div>
          <!-- Buttons Start -->
          <xsl:if test="$includeButtons = 'true'">
            <div class="filter-buttons">
              <button onclick="filterTests('all')">All</button>
              <button onclick="filterTests('Success')">Passed</button>
              <button onclick="filterTests('Failure')">Failed</button>
              <button onclick="filterTests('NotTested')">Not Tested</button>
            </div>
          </xsl:if>
          <!-- Buttons End -->
          <table>
            <tr>
              <th>Test Name</th>
              <th>Status</th>
              <th style="white-space: nowrap;">Failed Message</th>
            </tr>
            <xsl:for-each select="//test-case">
              <tr class="test-row">
                <xsl:attribute name="class">
                  <xsl:value-of select="concat('test-row ', @result)"/>
                </xsl:attribute>
                <td><xsl:value-of select="@name"/></td>
                <td>
                  <xsl:attribute name="class">
                    <xsl:choose>
                      <xsl:when test="@result='Success'">Success</xsl:when>
                      <xsl:when test="@result='Failure'">Failure</xsl:when>
                      <xsl:otherwise>not-tested</xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                  <xsl:value-of select="@result"/>
                </td>
                <td class="failure-message">
                  <xsl:if test="@result='Failure'">
                    <xsl:value-of select="failure/message"/>
                  </xsl:if>
                </td>
              </tr>
            </xsl:for-each>
          </table>
        </div>
      </body>
    </html>
  </xsl:template>
</xsl:stylesheet>
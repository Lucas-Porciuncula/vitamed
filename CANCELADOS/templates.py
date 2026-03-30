HTML_ZIMBRA = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin:0;padding:0;background-color:#f0f4f8;font-family:'Segoe UI',Arial,sans-serif;">

<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0f4f8;padding:30px 0;">
<tr><td align="center">

<table width="640" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.10);">

  <!-- HEADER -->
  <tr>
    <td style="background:linear-gradient(135deg,#1a3c5e 0%,#2e6da4 100%);padding:28px 32px;">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td>
            <span style="font-size:20px;font-weight:700;color:#ffffff;letter-spacing:1px;">VitalMed</span>
          </td>
          <td align="right">
            <span style="background:rgba(255,255,255,0.15);border-radius:20px;padding:4px 14px;font-size:12px;color:#ffffff;">
              📊 Relatório Diário
            </span>
          </td>
        </tr>
        <tr>
          <td colspan="2" style="padding-top:6px;">
            <span style="font-size:15px;color:rgba(255,255,255,0.85);">
              🚫 Relatório de Cancelamentos &nbsp;•&nbsp; {{DATA}}
            </span>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- KPI PRINCIPAL -->
  <tr>
    <td style="padding:24px 28px 0 28px;">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td style="background:#f5f8fb;border-radius:10px;padding:20px;text-align:center;">
            <p style="margin:0;font-size:11px;color:#7a8fa6;text-transform:uppercase;letter-spacing:0.6px;font-weight:600;">
              TOTAL DE CANCELAMENTOS
            </p>
            <p style="margin:8px 0 0;font-size:36px;font-weight:700;color:#1a3c5e;line-height:1;">
              {{TOTAL_CANCELADOS}}
            </p>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- KPIs ADICIONAIS -->
  <tr>
    <td style="padding:16px 28px 0 28px;">
      {{KPIS}}
    </td>
  </tr>

  <!-- DIVISOR -->
  <tr>
    <td style="padding:24px 28px 0 28px;">
      <hr style="border:none;border-top:1px solid #e8edf2;margin:0;">
    </td>
  </tr>

  <!-- TABELAS -->
  <tr>
    <td style="padding:20px 28px 28px 28px;">
      {{TABELAS}}
    </td>
  </tr>

  <!-- FOOTER -->
  <tr>
    <td style="background:#f5f8fb;padding:14px 28px;text-align:center;font-size:11px;color:#aab4bf;border-top:1px solid #e8edf2;">
      Relatório gerado automaticamente &nbsp;•&nbsp; Analytics VitalMed
    </td>
  </tr>

</table>
</td></tr>
</table>

</body>
</html>
"""


HTML_REATIVADOS = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin:0;padding:0;background-color:#f0f4f8;font-family:'Segoe UI',Arial,sans-serif;">

<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0f4f8;padding:30px 0;">
<tr><td align="center">

<table width="640" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.10);">

  <!-- HEADER -->
  <tr>
    <td style="background:linear-gradient(135deg,#0d4a2f 0%,#198754 100%);padding:28px 32px;">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td>
            <span style="font-size:20px;font-weight:700;color:#ffffff;letter-spacing:1px;">VitalMed</span>
          </td>
          <td align="right">
            <span style="background:rgba(255,255,255,0.15);border-radius:20px;padding:4px 14px;font-size:12px;color:#ffffff;">
              🔄 Relatório Diário
            </span>
          </td>
        </tr>
        <tr>
          <td colspan="2" style="padding-top:6px;">
            <span style="font-size:15px;color:rgba(255,255,255,0.85);">
              🔄 Relatório de Reativações &nbsp;•&nbsp; {{DATA}}
            </span>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- KPI PRINCIPAL -->
  <tr>
    <td style="padding:24px 28px 0 28px;">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td style="background:#f0faf4;border-radius:10px;padding:20px;text-align:center;">
            <p style="margin:0;font-size:11px;color:#7a8fa6;text-transform:uppercase;letter-spacing:0.6px;font-weight:600;">
              TOTAL DE REATIVADOS
            </p>
            <p style="margin:8px 0 0;font-size:36px;font-weight:700;color:#198754;line-height:1;">
              {{TOTAL_REATIVADOS}}
            </p>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- KPIs ADICIONAIS -->
  <tr>
    <td style="padding:16px 28px 0 28px;">
      {{KPIS}}
    </td>
  </tr>

  <!-- DIVISOR -->
  <tr>
    <td style="padding:24px 28px 0 28px;">
      <hr style="border:none;border-top:1px solid #e8edf2;margin:0;">
    </td>
  </tr>

  <!-- TABELAS -->
  <tr>
    <td style="padding:20px 28px 28px 28px;">
      <p style="margin:0 0 16px;font-size:13px;font-weight:600;color:#198754;text-transform:uppercase;letter-spacing:0.6px;">
        Lista Detalhada de Reativações
      </p>
      {{TABELAS}}
    </td>
  </tr>

  <!-- FOOTER -->
  <tr>
    <td style="background:#f5f8fb;padding:14px 28px;text-align:center;font-size:11px;color:#aab4bf;border-top:1px solid #e8edf2;">
      Relatório de Recuperação de Clientes &nbsp;•&nbsp; Analytics VitalMed
    </td>
  </tr>

</table>
</td></tr>
</table>

</body>
</html>
"""
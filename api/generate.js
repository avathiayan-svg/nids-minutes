const axios = require('axios');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        WidthType, BorderStyle, ShadingType, AlignmentType } = require('docx');

async function getShiftCareNotes(email, password, clientId) {
  const loginPageResp = await axios.get('https://app.shiftcare.com/users/sign_in', {
    headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'text/html' },
    maxRedirects: 5,
  });
  const csrfMatch = loginPageResp.data.match(/name="authenticity_token"[^>]*value="([^"]+)"/);
  const csrf = csrfMatch ? csrfMatch[1] : '';
  const cookies = loginPageResp.headers['set-cookie'] || [];
  const cookieStr = cookies.map(c => c.split(';')[0]).join('; ');

  const loginResp = await axios.post('https://app.shiftcare.com/users/sign_in',
    `user[email]=${encodeURIComponent(email)}&user[password]=${encodeURIComponent(password)}&authenticity_token=${encodeURIComponent(csrf)}`,
    {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Cookie': cookieStr,
        'User-Agent': 'Mozilla/5.0',
        'Referer': 'https://app.shiftcare.com/users/sign_in',
      },
      maxRedirects: 5,
    }
  );

  const sessionCookies = loginResp.headers['set-cookie'] || [];
  const allCookies = [...cookies, ...sessionCookies]
    .map(c => c.split(';')[0])
    .filter((v, i, a) => a.findIndex(c => c.split('=')[0] === v.split('=')[0]) === i)
    .join('; ');

  const notesResp = await axios.get(
    `https://app.shiftcare.com/users/clients/${clientId}?tab=communication`,
    {
      headers: {
        'Cookie': allCookies,
        'User-Agent': 'Mozilla/5.0',
        'Accept': 'text/html',
      },
      maxRedirects: 5,
    }
  );

  const html = notesResp.data;
  const text = html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const start = text.indexOf('Add Notes');
  return start > -1 ? text.slice(start, start + 8000) : text.slice(0, 8000);
}

async function fillTemplate(notes, client, meetingDate, nextDate, anthropicKey) {
  const prompt = `You are filling a Weekly Staff Meeting Minutes form for an NDIS disability support provider in Perth, Australia. Client: ${client}.

ShiftCare shift notes:
${notes.slice(0, 5000)}

Return ONLY valid JSON, no explanation, no markdown:
{
  "meeting_date": "${meetingDate}",
  "attendance": [
    {"name":"Shanti","role":"Director","status":"P"},
    {"name":"Alan","role":"Director","status":"P"},
    {"name":"Jocelyn","role":"Support Worker","status":"P"},
    {"name":"Christy","role":"Support Worker","status":"P"}
  ],
  "prev_actions": [{"action":"","responsible":"","status":"","comments":""}],
  "client_update": [{"area":"","observations":"","responsible":""}],
  "appointments": [{"date":"","activity":"","responsible":"","notes":""}],
  "communication": {"support_coordinator":false,"public_trustee":false,"family_updates":false,"other":"","notes":""},
  "issues": [{"issue":"","discussion":"","action":"","responsible":""}],
  "next_date": "${nextDate}",
  "next_time": "12:00 PM"
}

Use areas: Personal Care, Health & Wellbeing, Medication, Nutrition, Daily Living, Community Access, Dialysis, Transport, Social/Emotional.
Only include rows with real content. Flag issues only if genuinely concerning.`;

  const resp = await axios.post('https://api.anthropic.com/v1/messages', {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 1500,
    messages: [{ role: 'user', content: prompt }]
  }, {
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': anthropicKey,
      'anthropic-version': '2023-06-01'
    }
  });

  const raw = resp.data.content.map(b => b.text || '').join('');
  const clean = raw.replace(/```json|```/g, '').trim();
  return JSON.parse(clean);
}

async function buildDocx(data, client) {
  const NAVY='1F4E79', WHITE='FFFFFF', LGREY='F2F2F2';
  const sb={style:BorderStyle.SINGLE,size:4,color:'AAAAAA'};
  const B={top:sb,bottom:sb,left:sb,right:sb};
  function hc(t,w){return new TableCell({borders:B,width:{size:w,type:WidthType.DXA},shading:{fill:NAVY,type:ShadingType.CLEAR},margins:{top:70,bottom:70,left:110,right:110},children:[new Paragraph({children:[new TextRun({text:String(t||''),bold:true,color:WHITE,size:19,font:'Arial'})]})]})}
  function dc(t,w,s){return new TableCell({borders:B,width:{size:w,type:WidthType.DXA},shading:{fill:s||WHITE,type:ShadingType.CLEAR},margins:{top:70,bottom:70,left:110,right:110},children:[new Paragraph({children:[new TextRun({text:String(t||''),size:19,font:'Arial'})]})]})}
  function sh(t){return new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],rows:[new TableRow({children:[new TableCell({borders:B,width:{size:9360,type:WidthType.DXA},shading:{fill:NAVY,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:140,right:140},children:[new Paragraph({children:[new TextRun({text:t,bold:true,color:WHITE,size:20,font:'Arial'})]})]})]})]})}
  function gap(){return new Paragraph({spacing:{before:120,after:60},children:[new TextRun('')]})}
  function sgap(){return new Paragraph({spacing:{before:60,after:40},children:[new TextRun('')]})}
  function mkt(cols,hdrs,rows,min){
    const tot=cols.reduce((a,b)=>a+b,0);
    let d=rows&&rows.length?rows:[];
    while(d.length<(min||0))d.push(new Array(cols.length).fill(''));
    return new Table({width:{size:tot,type:WidthType.DXA},columnWidths:cols,rows:[new TableRow({children:hdrs.map((h,i)=>hc(h,cols[i]))}), ...d.map((r,ri)=>new TableRow({children:cols.map((w,ci)=>dc(r[ci]||'',w,ri%2===0?WHITE:LGREY))}))]});
  }
  function tick(checked,label){return new TableRow({children:[new TableCell({borders:B,width:{size:9360,type:WidthType.DXA},margins:{top:70,bottom:70,left:140,right:140},children:[new Paragraph({children:[new TextRun({text:(checked?'☒':'☐')+'  '+label,size:19,font:'Arial'})]})]})]});}
  const att=mkt([3120,3120,3120],['Name','Role','Present (P) / Absent (A)'],(data.attendance||[]).map(a=>[a.name,a.role,`${a.name} (${a.status})`]),4);
  const prev=mkt([2880,1800,1800,2880],['Action Item','Responsible Staff','Status / Outcome','Comments'],(data.prev_actions||[]).filter(r=>r.action).map(r=>[r.action,r.responsible,r.status,r.comments]),2);
  const cu=mkt([2000,4760,2600],['Area','Observations / Updates','Responsible Staff / Notes'],(data.client_update||[]).filter(r=>r.area).map(r=>[r.area,r.observations,r.responsible]),9);
  const ap=mkt([1800,3000,2160,2400],['Date','Appointment / Activity','Responsible Staff','Notes'],(data.appointments||[]).filter(r=>r.activity).map(r=>[r.date,r.activity,r.responsible,r.notes]),6);
  const comm=data.communication||{};
  const ct=new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],rows:[tick(comm.support_coordinator,'Contact with Support Coordinator required'),tick(comm.public_trustee,'Contact with Public Trustee required'),tick(comm.family_updates,'Family communication updates'),tick(!!comm.other,'Other (specify): '+(comm.other||''))]});
  const iss=mkt([2000,2800,2400,2160],['Issue','Discussion Summary','Action Required','Responsible Person'],(data.issues||[]).filter(r=>r.issue).map(r=>[r.issue,r.discussion,r.action,r.responsible]),3);
  const nt=new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[2000,7360],rows:[new TableRow({children:[hc('Date',2000),dc(data.next_date||'',7360)]}),new TableRow({children:[hc('Time',2000),dc(data.next_time||'',7360,LGREY)]})]});
  const sig=mkt([3120,3120,3120],['Staff Name','Signature','Date'],['Shanti','Alan','Jocelyn','Christy'].map(n=>[n,'','']),0);
  const doc=new Document({sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1080,right:1080,bottom:1080,left:1080}}},children:[
    new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:80},children:[new TextRun({text:`Weekly Staff Meeting Minutes – ${client}`,bold:true,size:30,font:'Arial',color:NAVY})]}),
    new Paragraph({spacing:{before:60,after:40},children:[new TextRun({text:`Meeting Date: ${data.meeting_date}`,bold:true,size:20,font:'Arial'})]}),
    new Paragraph({spacing:{before:0,after:180},children:[new TextRun({text:'Time: 12:00 PM',bold:true,size:20,font:'Arial'})]}),
    sh('1. Attendance'),sgap(),att,gap(),
    sh("2. Review of Previous Week's Notes / Actions"),sgap(),prev,gap(),
    sh(`3. Client Update – ${client}`),sgap(),cu,gap(),
    sh('4. Upcoming Appointments and Plans'),sgap(),ap,gap(),
    sh('5. Communication and Coordination'),sgap(),ct,
    new Paragraph({spacing:{before:80,after:30},children:[new TextRun({text:'Notes:',bold:true,size:19,font:'Arial'})]}),
    new Paragraph({spacing:{before:0,after:180},children:[new TextRun({text:comm.notes||'',size:19,font:'Arial'})]}),
    sh('6. Issues or Concerns Raised'),sgap(),iss,gap(),
    sh('7. Next Meeting'),sgap(),nt,gap(),
    sh('8. Signatures'),sgap(),sig,
  ]}]});
  return await Packer.toBuffer(doc);
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });
  try {
    const { meetingDate, nextDate } = req.body;
    const scEmail = process.env.SHIFTCARE_EMAIL;
    const scPassword = process.env.SHIFTCARE_PASSWORD;
    const anthropicKey = process.env.ANTHROPIC_API_KEY;
    if (!scEmail || !scPassword || !anthropicKey) {
      return res.status(500).json({ error: 'Server not configured. Please contact your administrator.' });
    }
    const [santoshNotes, thirzaNotes] = await Promise.all([
      getShiftCareNotes(scEmail, scPassword, '921417'),
      getShiftCareNotes(scEmail, scPassword, '1054663'),
    ]);
    const [santoshData, thirzaData] = await Promise.all([
      fillTemplate(santoshNotes, 'Santosh Mon Abraham', meetingDate, nextDate, anthropicKey),
      fillTemplate(thirzaNotes, 'Thirza Stirling', meetingDate, nextDate, anthropicKey),
    ]);
    const [santoshBuf, thirzaBuf] = await Promise.all([
      buildDocx(santoshData, 'Santosh Mon Abraham'),
      buildDocx(thirzaData, 'Thirza Stirling'),
    ]);
    res.status(200).json({
      santosh: santoshBuf.toString('base64'),
      thirza: thirzaBuf.toString('base64'),
      santoshName: `Minutes_Santosh_${meetingDate.replace(/\//g,'')}.docx`,
      thirzaName: `Minutes_Thirza_${meetingDate.replace(/\//g,'')}.docx`,
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || 'Something went wrong. Please try again.' });
  }
};

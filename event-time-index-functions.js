/*******************************************************
 * EVENT TIME FIX — Index.html
 *
 * In the live Index.html, replace the existing formatTimeInputValue
 * and formatTime12 functions with the two functions below.
 *
 * formatTimeInputValue keeps feeding <input type="time"> (HH:mm).
 * formatTime12 shows the saved 12-hour text in Overview.
 *******************************************************/

function formatTimeInputValue(value){
  if(!value) return '';
  const s=String(value).trim().replace(/^'/,'').trim();
  if(!s) return '';

  // Check 12-hour first so "11:30 PM" is not parsed as 11:30 (AM).
  let m=s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)$/i);
  if(m){
    let h=Number(m[1]), min=m[2], ap=m[3].toUpperCase();
    if(ap==='AM'&&h===12)h=0;
    if(ap==='PM'&&h!==12)h+=12;
    return String(h).padStart(2,'0')+':'+min;
  }

  m=s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if(m) return String(Number(m[1])).padStart(2,'0')+':'+m[2];

  if(/\d{4}-\d{2}-\d{2}[T ]/.test(s)){
    const d=new Date(s);
    if(!isNaN(d.getTime())){
      return String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0');
    }
  }
  return '';
}

function formatTime12(value){
  if(!value) return '';
  const s=String(value).trim().replace(/^'/,'').trim();
  if(!s) return '';

  // Already saved 12-hour text — show it as saved.
  let m=s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)$/i);
  if(m){
    let h=Number(m[1]);
    const min=m[2];
    const ap=m[3].toUpperCase();
    if(h<1) h=12;
    if(h>12) h=((h-1)%12)+1;
    return h+':'+min+' '+ap;
  }

  // Legacy 24-hour HH:mm / HH:mm:ss
  m=s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if(m){
    let h=Number(m[1]);
    const min=m[2];
    const ap=h>=12?'PM':'AM';
    h=h%12||12;
    return h+':'+min+' '+ap;
  }

  // Sheets day-fraction
  if(/^0(?:\.\d+)?$/.test(s) || /^\d*\.\d+$/.test(s)){
    const fraction=Number(s);
    if(isFinite(fraction)){
      const totalMinutes=Math.round((fraction-Math.floor(fraction))*24*60);
      let h=Math.floor(totalMinutes/60)%24;
      const min=String(totalMinutes%60).padStart(2,'0');
      const ap=h>=12?'PM':'AM';
      h=h%12||12;
      return h+':'+min+' '+ap;
    }
  }

  // ISO / 1899 Sheets serial — use getHours(), not toLocaleTimeString()
  if(/\d{4}-\d{2}-\d{2}[T ]/.test(s)){
    const d=new Date(s);
    if(!isNaN(d.getTime())){
      let h=d.getHours();
      const min=String(d.getMinutes()).padStart(2,'0');
      const ap=h>=12?'PM':'AM';
      h=h%12||12;
      return h+':'+min+' '+ap;
    }
  }

  return s;
}

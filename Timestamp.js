/**
 * Notes:
 * * https://developers.google.com/apps-script/
 * * https://developers.google.com/apps-script/concepts/scopes
 * * https://developers.google.com/apps-script/guides/clasp
 * * https://developers.google.com/gsuite/add-ons/editors/docs/quickstart/translate
 * * https://developers.google.com/gsuite/add-ons/how-tos/publish-addons
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Insert Timestamp', 'insertTimestamp')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

/**
 * What makes this difficult:
 * - JavaScript doesn't have a strftime
 */
function insertTimestamp() {
  const date = new Date();
  const fmt = "%Y/%m/%d %H:%M ";
  insertText(strftime(fmt, date));
}

/**
 * Replaces the text of the current selection with the provided text, or
 * inserts text at the current cursor location. (There will always be either
 * a selection or a cursor.) If multiple elements are selected, only inserts the
 * translated text in the first element that can contain text and removes the
 * other elements.
 *
 * @param {string} newText The text with which to replace the current selection.
 */
function insertText(newText) {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    var replaced = false;
    var elements = selection.getSelectedElements();
    if (elements.length === 1 && elements[0].getElement().getType() ===
        DocumentApp.ElementType.INLINE_IMAGE) {
      throw new Error('Can\'t insert text into an image.');
    }
    for (var i = 0; i < elements.length; ++i) {
      if (elements[i].isPartial()) {
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();
        element.deleteText(startIndex, endIndex);
        if (!replaced) {
          element.insertText(startIndex, newText);
          replaced = true;
        } else {
          // This block handles a selection that ends with a partial element. We
          // want to copy this partial text to the previous element so we don't
          // have a line-break before the last partial.
          var parent = element.getParent();
          var remainingText = element.getText().substring(endIndex + 1);
          parent.getPreviousSibling().asText().appendText(remainingText);
          // We cannot remove the last paragraph of a doc. If this is the case,
          // just remove the text within the last paragraph instead.
          if (parent.getNextSibling()) {
            parent.removeFromParent();
          } else {
            element.removeFromParent();
          }
        }
      } else {
        var element = elements[i].getElement();
        if (!replaced && element.editAsText) {
          // Only translate elements that can be edited as text, removing other
          // elements.
          element.clear();
          element.asText().setText(newText);
          replaced = true;
        } else {
          // We cannot remove the last paragraph of a doc. If this is the case,
          // just clear the element.
          if (element.getNextSibling()) {
            element.removeFromParent();
          } else {
            element.clear();
          }
        }
      }
    }
  } else {
    var cursor = DocumentApp.getActiveDocument().getCursor();
    var surroundingText = cursor.getSurroundingText().getText();
    var surroundingTextOffset = cursor.getSurroundingTextOffset();

    // If the cursor follows or preceds a non-space character, insert a space
    // between the character and the translation. Otherwise, just insert the
    // translation.
    if (surroundingTextOffset > 0) {
      if (surroundingText.charAt(surroundingTextOffset - 1) != ' ') {
        newText = ' ' + newText;
      }
    }
    if (surroundingTextOffset < surroundingText.length) {
      if (surroundingText.charAt(surroundingTextOffset) != ' ') {
        newText += ' ';
      }
    }
    const el = cursor.insertText(newText);
    const doc = DocumentApp.getActiveDocument()
    const position = doc.newPosition(el, newText.length);
//    doc.setCursor(doc.newPosition(el.getParent().getNextSibling()));
    doc.setCursor(position);
  }
}
  
/* source of strftime
 * https://github.com/iapyeh/TextFactory/blob/42f601ccef8ca323e490d7048b4e75d0dbaab068/source/Code.gs#L915
 */
function thousandCommas(n,sp) {
  if (!sp) sp = ','
  return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, sp);
}
function zeropad(n, size) {
    n = '' + n; /* Make sure it's a string */
    size = size || 2;
    while (n.length < size) {
        n = '0' + n;
    }
    return n;
}
function getWeek(d){
    var onejan = new Date(d.getFullYear(), 0, 1);
    return Math.ceil(((d - onejan)/86400000 - (6-onejan.getDay())%6)/7);
}
function getWeekMonday(d){
    var onejan = new Date(d.getFullYear(), 0, 1);
    return 1+Math.floor(((d - onejan)/86400000 - ((onejan.getDay()+6)%7 + 1))/7);
}
function getDayOfYear(d){
    var j1= new Date(d);
    j1.setMonth(0, 0);
    return Math.round((d-j1)/8.64e7);
}
function twelve(n) {
    return (n <= 12) ? (n==0 ? 12 : n) : n - 12;
}
function tzOffset(offset){
  var s = ((offset<0? '+':'-')+ // Note the reversed sign!
          zeropad(my_parseInt(Math.abs(offset/60)), 2)+
          zeropad(Math.abs(offset%60), 2))
  return s
}

function strftime(format, date,loc) {
    date = date || getNewDate();
    var locfmtpat = /\%[abAB]/; 
    if (locfmtpat.test(format) && !locale[loc]){
      initializeUserLocale();
    }
    var l = locale[loc] || locale['en'];
    var months = l.months, days = l.days;
    var hd; //hebrew data
    var fields = {
        a: l.daysshort[date.getDay()],
        A: days[date.getDay()],
        b: l.monthsshort[date.getMonth()],
        B: months[date.getMonth()],
//        c: date.toLocaleString(),
        d: zeropad(date.getDate()),
        H: zeropad(date.getHours()),
        h: zeropad(twelve(date.getHours())),
        j: getDayOfYear(date),
        m: zeropad(date.getMonth() + 1),
        M: zeropad(date.getMinutes()),
        n: date.getMonth() + 1,
        N: date.getDate(),
        p: (date.getHours() >= 12) ? 'PM' : 'AM',
        S: zeropad(date.getSeconds()),
        w: zeropad(date.getDay() + 1),
        W: getWeekMonday(date)+1,
        U: getWeek(date)+1,
//        x: date.toLocaleDateString(),
//        X: date.toLocaleTimeString(),
        y: ('' + date.getFullYear()).substr(2, 4),
        Y: '' + date.getFullYear(),
        Z: tzOffset(date.getTimezoneOffset()),
        '%' : '%',
    };
    var result = '', i = 0, len=format.length;
    var hd = {hy:0,hm:'-',hd:'-',events:[],hebrew:'-'}
    var lu = {year:0,month:'',day:''}
    while (i < format.length) {
        if (format[i] === '%' && (i+3<len) && (format[i+1]=='H') && (format[i+2]=='e')) {
            if (hd.hy==0) {
              var key = date.getFullYear()+':'+date.getMonth()+':'+date.getDate()
              if (hebrew_date_cache[key]) hd = hebrew_date_cache[key]
              else {
                hd.hy = '2016'
                getHebrew(date,key,function(result){
                   hd = result
                })
              }
            }
            // %HeY, %HeM, %HeD
            switch (format[i + 3]){
              case 'Y':
                  result = result + hd.hy
                  break
              case 'M':
                  result = result + hd.hm
                  break
              case 'D':
                  result = result + hd.hd
                  break
              case 'H':
                  result = result + hd.hebrew
                  break
              case 'E':
                  result = result + (hd.events ? hd.events.join(', ') : '')
                  break
              default:
                  result = result + format[i]+format[i + 1]+format[i + 2]+format[i + 3]
            }            
            i += 3;
        }  
        else if (format[i] === '%' && (i+3<len) && (format[i+1]=='L') && (format[i+2]=='u')) {
            if (lu.year==0) {
              var key = date.getFullYear()+':'+date.getMonth()+':'+date.getDate()
              if (lunar_date_cache[key]) {
                lu = lunar_date_cache[key]
              }
              else {
                lu = getLunar(date,key)
              }
            }
            // %LuY, %LuM, %LuD
            switch (format[i + 3]){
              case 'Y':
                  result = result + lu.year
                  break
              case 'M':
                  result = result + lu.month
                  break
              case 'D':
                  result = result + lu.day
                  break
              default:
                  result = result + format[i]+format[i + 1]+format[i + 2]+format[i + 3]
            }            
            i += 3;
        }
        else if (format[i] === '%' && (i+3<len) && (format[i+1]=='e') && (format[i+2]=='n') && (format[i+3]=='b'||format[i+3]=='B')) {        
          // %enB, %enB
            result = result + (format[i+3]=='b' ? locale.en.monthsshort[date.getMonth()] : locale.en.months[date.getMonth()])
            i += 3;          
        }
        else if ((format[i] == '%') && (i+2<len) && (format[i+1]=='*')) {
            // %*d, %*H, %*S, 
            result = result + my_parseInt(fields[format[i + 2]]);
            i+=2;
        }
        else if ((format[i] == '%') && (i+2<len) && (format[i+1]=='+')) {
            // %+d
            var n = my_parseInt(fields[format[i + 2]])
            if (n>=11 && n<=13){
              n = n+'th'
            }
            else if (n%10==1) n = n+'st'
            else if (n%10==2) n = n+'nd'
            else if (n%10==3) n = n+'rd'
            else n = n+'th'
            result = result + n;
            i += 2;
        }        
        else if (format[i] === '%' && (i+1<len)) {
            // one char format, ex %Y, %M
            result = result + fields[format[i + 1]];
            ++i;
        }
        else {
            // regular char, not in format
            result = result + format[i];
        }
        ++i;
    }
    return result;
}

var locale={'en':{
  days:["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"],
  daysshort:["Sun","Mon","Tue","Wed","Thu","Fri","Sat"],
  months:["January","February","March","April","May","June","July","August","September","October","November","December"],
  monthsshort:["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
  }
}

function my_parseInt(s){
  // correct convert 08,09 to 8 and 9
  //bug of google: https://issuetracker.google.com/issues/36759856
   return typeof(s)=='string' ? (s.indexOf('.')>0 ? parseFloat(s) : parseInt(s.replace(/^0+/,''))) : s
}

var hebrew_date_cache={};
function getHebrew(d,key,callback){
  //https://www.hebcal.com/home/219/hebrew-date-converter-rest-api
  //var d = new Date()
  var url = 'http://www.hebcal.com/converter/?cfg=json&gy='+d.getFullYear()+'&gm='+(1+d.getMonth())+'&gd='+d.getDate()+'&g2h=1'
  try{
    var hd = UrlFetchApp.fetch(url).getContentText()
    var obj = JSON.parse(hd)
    hebrew_date_cache[key] = obj
    callback(obj)
  }
  catch(e){
    Logger.log(e)
    callback(null)
  }
}

function LunarCalendar(){
    this.lunarInfo=new Array(  
        0x04bd8,0x04ae0,0x0a570,0x054d5,0x0d260,0x0d950,0x16554,0x056a0,0x09ad0,0x055d2,  
        0x04ae0,0x0a5b6,0x0a4d0,0x0d250,0x1d255,0x0b540,0x0d6a0,0x0ada2,0x095b0,0x14977,  
        0x04970,0x0a4b0,0x0b4b5,0x06a50,0x06d40,0x1ab54,0x02b60,0x09570,0x052f2,0x04970,  
        0x06566,0x0d4a0,0x0ea50,0x06e95,0x05ad0,0x02b60,0x186e3,0x092e0,0x1c8d7,0x0c950,  
        0x0d4a0,0x1d8a6,0x0b550,0x056a0,0x1a5b4,0x025d0,0x092d0,0x0d2b2,0x0a950,0x0b557,  
        0x06ca0,0x0b550,0x15355,0x04da0,0x0a5b0,0x14573,0x052b0,0x0a9a8,0x0e950,0x06aa0,  
        0x0aea6,0x0ab50,0x04b60,0x0aae4,0x0a570,0x05260,0x0f263,0x0d950,0x05b57,0x056a0,  
        0x096d0,0x04dd5,0x04ad0,0x0a4d0,0x0d4d4,0x0d250,0x0d558,0x0b540,0x0b6a0,0x195a6,  
        0x095b0,0x049b0,0x0a974,0x0a4b0,0x0b27a,0x06a50,0x06d40,0x0af46,0x0ab60,0x09570,  
        0x04af5,0x04970,0x064b0,0x074a3,0x0ea50,0x06b58,0x055c0,0x0ab60,0x096d5,0x092e0,  
        0x0c960,0x0d954,0x0d4a0,0x0da50,0x07552,0x056a0,0x0abb7,0x025d0,0x092d0,0x0cab5,  
        0x0a950,0x0b4a0,0x0baa4,0x0ad50,0x055d9,0x04ba0,0x0a5b0,0x15176,0x052b0,0x0a930,  
        0x07954,0x06aa0,0x0ad50,0x05b52,0x04b60,0x0a6e6,0x0a4e0,0x0d260,0x0ea65,0x0d530,  
        0x05aa0,0x076a3,0x096d0,0x04bd7,0x04ad0,0x0a4d0,0x1d0b6,0x0d250,0x0d520,0x0dd45,  
        0x0b5a0,0x056d0,0x055b2,0x049b0,0x0a577,0x0a4b0,0x0aa50,0x1b255,0x06d20,0x0ada0,  
        0x14b63); 
        
    this.Gan=new Array("甲","乙","丙","丁","戊","己","庚","辛","壬","癸");  
    this.Zhi=new Array("子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥");  
    this.nStr1 = new Array('','一','二','三','四','五','六','七','八','九','十');  
    this.nStr2 = new Array('初','十','廿','卅','□');
}
LunarCalendar.prototype = {
    lYearDays : function(y) {  
        var i, sum = 348;  
        for(i=0x8000; i>0x8; i>>=1) sum += (this.lunarInfo[y-1900] & i)? 1: 0;  
        return(sum+this.leapDays(y));  
    },
    leapDays : function(y) {  
        if(this.leapMonth(y))  return((this.lunarInfo[y-1900] & 0x10000)? 30: 29);  
        else return(0);  
    },
    leapMonth : function(y) {  
        return(this.lunarInfo[y-1900] & 0xf);  
    },
    monthDays : function (y,m) {  
        return( (this.lunarInfo[y-1900] & (0x10000>>m))? 30: 29 );  
    },
    md2Str:function(n,month){
      var n1 = Math.floor(n/10)
      var n0 = n - (n1*10)
      if (month) return (n1==0 ? '' : this.nStr2[n1])+(n0==0 ? '' : this.nStr1[n0])
      else return (this.nStr2[n1])+(n0==0 ? '' : this.nStr1[n0])
    },
    toLunar: function (objDate) {  

        var lunarDate = {}
      
        var i, leap=0, temp=0;  
        var offset   = (Date.UTC(objDate.getFullYear(),objDate.getMonth(),objDate.getDate()) - Date.UTC(1900,0,31))/86400000;  
  
        for(i=1900; i<2050 && offset>0; i++) { temp=this.lYearDays(i); offset-=temp; }  
  
        if(offset<0) { offset+=temp; i--; }  
  
        lunarDate.year = i;  
  
        leap = this.leapMonth(i); //闰哪个月  
        lunarDate.isLeap = false;  
  
        for(i=1; i<13 && offset>0; i++) {  
            //闰月  
            if(leap>0 && i==(leap+1) && lunarDate.isLeap==false)  
                { --i; lunarDate.isLeap = true; temp = this.leapDays(lunarDate.year); }  
            else  
                { temp = this.monthDays(lunarDate.year, i); }  
  
            //解除闰月  
            if(lunarDate.isLeap==true && i==(leap+1)) lunarDate.isLeap = false;  
  
            offset -= temp;  
        }  
  
        if(offset==0 && leap>0 && i==leap+1)  
        if(lunarDate.isLeap)  
            { tlunarDatehis.isLeap = false; }  
        else  
            { lunarDate.isLeap = true; --i; }  
  
        if(offset<0){ offset += temp; --i; }  
  
        lunarDate.year = this.cyclical(lunarDate.year)
        lunarDate.month = this.md2Str(i,true)
        lunarDate.day = this.md2Str(offset + 1);
        return lunarDate
    },
    cyclical:function(year) {  
        var num = year - 1900 + 36
        return(this.Gan[num%10]+this.Zhi[num%12]);  
    }
}

var lunar_date_cache = {}
var lunar_calendar;
function getLunar(d,key){
  if (!lunar_calendar) lunar_calendar = new LunarCalendar()
  var lu = lunar_calendar.toLunar(d)
  lunar_date_cache[key] = lu
  return lu
}
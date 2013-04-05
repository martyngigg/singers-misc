/**
 * Returns the date of the next day given by dayToFindIndex (Sunday=0) after
 * 9pm from the given start date
 */
function getNextDate(start,dayToFindIndex)
{
  var dowOffsets = new Array();
  for(var i=0;i<7;++i)
  {
    var offset = (dayToFindIndex - i);
    if(offset < 0) offset += 7;
    dowOffsets[i] = offset;
  }
  // If we are on the actual day it must be after 9pm to be next week
  if(start.getHours() > 21) dowOffsets[dayToFindIndex] = 7;
  var dayToFind = new Date();
  dayToFind.setDate(start.getDate() + dowOffsets[start.getDay()]);
  return dayToFind;
};
/**
 * Add endsWith function to String class
 */
String.prototype.endsWith = function(suffix) {
    return this.indexOf(suffix, this.length - suffix.length) !== -1;
};

/**
 * Returns the superscript for the given date
 */
function getSuperscript(dateAsStr)
{
  var nchars=dateAsStr.length; // 1 or 2
  var superscript = "";
  if(nchars==1 || (nchars == 2 && dateAsStr.charAt(0) != "1")) {
    if(dateAsStr.endsWith("1")) superscript = "st";
    else if(dateAsStr.endsWith("2")) superscript = "nd";
    else if(dateAsStr.endsWith("3")) superscript = "rd";
    else superscript = "th";
  }
  else {
    superscript = "th";
  }
  return superscript;
}

/**
 * Returns the date of the next Thursday from the current date. Uses getNextDate()
 */
function nextThursdayAsHTML()
{
  var dayToFindIndex=4; // Thursday
  var now = new Date();
  nextThursday = getNextDate(now,dayToFindIndex);
  var monthNames = [ "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December" ];
  var dateAsStr =  nextThursday.getDate().toString(); // This is the numbered day of the month (1-31)
  var superscript = getSuperscript(dateAsStr)
  return dateAsStr + "<sup>" + superscript + "</sup> " + monthNames[nextThursday.getMonth()]
};

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
 * Returns the date of the next Thursday from the current date. Uses getNextDate()
 */
function nextThursdayAsHTML()
{
  var dayToFindIndex=4; // Thursday
  var now = new Date();
  nextThursday = getNextDate(now,dayToFindIndex);
  var monthNames = [ "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December" ];
  var dateAsStr =  nextThursday.getDate().toString();
  var superscript = "";
  if(dateAsStr.endsWith("1")) superscript = "st";
  else if(dateAsStr.endsWith("2")) superscript = "nd";
  else if(dateAsStr.endsWith("3")) superscript = "rd";
  else superscript = "th";
  return dateAsStr + "<sup>" + superscript + "</sup> " + monthNames[nextThursday.getMonth()]
};

function ZG_parser_Pictures_Names(){
// со страницы https://protivogaz.ru/catalog/
// взять ссылки, содержащие https://protivogaz.ru/catalog/
// проходом рекурсивным по этим ссылкам создать массив Названий и ссылок
// если на странице отрабатывают 2 XPath
// массив на лист Картинки

const sURL_start = "https://protivogaz.ru/catalog/";


}

// исследовать 
// http://googleappscripting.com/working-with-urls/
// https://web.archive.org/web/20201121023356/http://googleappscripting.com/doget-dopost-tutorial-examples/

// образец для разбора https://www.sites.google.com/site/scriptsexamples/learn-by-example/parsing-html
function doGet() {

  var html = UrlFetchApp.fetch('http://en.wikipedia.org/wiki/Document_Object_Model').getContentText();

  var doc = XmlService.parse(html);

  var html = doc.getRootElement();

  var menu = getElementsByClassName(html, 'vertical-navbox nowraplinks')[0];

  var output = XmlService.getRawFormat().format(menu);

  return HtmlService.createHtmlOutput(output);

}  

// We fetch the HTML through UrlFetch

// We use the XMLService to parse this HTML

// Then we can use a specific function to grab the element we want in the DOM tree (like getElementsByClassName)

// And we convert back this element to HTML 

// Or we could get all the links / anchors available in this menu and display them

function doGet() {

  var html = UrlFetchApp.fetch('http://en.wikipedia.org/wiki/Document_Object_Model').getContentText();

  var doc = XmlService.parse(html);

  var html = doc.getRootElement();

  var menu = getElementsByClassName(html, 'vertical-navbox nowraplinks')[0];

  var output = '';

  var linksInMenu = getElementsByTagName(menu, 'a');

  for(i in linksInMenu) output+= XmlService.getRawFormat().format(linksInMenu[i])+'<br>';

  return HtmlService.createHtmlOutput(output);

}


function getElementById(element, idToFind) {  

  var descendants = element.getDescendants();  

  for(i in descendants) {

    var elt = descendants[i].asElement();

    if( elt !=null) {

      var id = elt.getAttribute('id');

      if( id !=null && id.getValue()== idToFind) return elt;    

    }

  }

}


function getElementsByClassName(element, classToFind) {  

  var data = [];

  var descendants = element.getDescendants();

  descendants.push(element);  

  for(i in descendants) {

    var elt = descendants[i].asElement();

    if(elt != null) {

      var classes = elt.getAttribute('class');

      if(classes != null) {

        classes = classes.getValue();

        if(classes == classToFind) data.push(elt);

        else {

          classes = classes.split(' ');

          for(j in classes) {

            if(classes[j] == classToFind) {

              data.push(elt);

              break;

            }

          }

        }

      }

    }

  }

  return data;

}

function getElementsByTagName(element, tagName) {  

  var data = [];

  var descendants = element.getDescendants();  

  for(i in descendants) {

    var elt = descendants[i].asElement();     

    if( elt !=null && elt.getName()== tagName) data.push(elt);      

  }

  return data;

}
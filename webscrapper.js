const reader = require('xlsx')
const file = reader.readFile('Input.xlsx');
const puppeteer = require('puppeteer');
const worksheet = file.Sheets["Sheet1"];
var wb = reader.utils.book_new();
const ISBN=reader.utils.sheet_to_json(worksheet, {
    raw: true,
    range: "C1:C6",
      defval: null,
  })
const Title= reader.utils.sheet_to_json(worksheet, {
    raw: true,
    range: "B1:B6",
      defval: null,
  })
  function match(s,t){
    let n=s.length;
    let m=t.length;
    s=s.toLowerCase();
    t=t.toLowerCase();
    let i=0;
    let count=0;
    for(i=0;i<m;i++){
      if(t[i]==s[i]){
        count+=1;
      }
    }
    let z=n*0.9;
    return count>=z;
  }

async function scrape() {
    try{
      const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();
    let i=0;
    for(i=0;i<ISBN.length;i++){
      await page.goto('https://www.snapdeal.com');
    
    
      await page.waitForSelector('#inputValEnter');
      await page.type('#inputValEnter', JSON.stringify(ISBN[i].ISBN));
      await page.click('.searchformButton');
      await page.waitForSelector('#products')
      let Books = await page.$$eval('#products section div.js-tuple', books => {
        
        let titles= books.map(el =>  el.querySelector('div.product-tuple-description div a > p').title);
      
        ids= books.map(el => el.querySelector('div.product-tuple-description div a').getAttribute('href'));
        let prices= books.map(el =>  parseInt(el.querySelector('div.product-tuple-description div a > div > div span.product-price').getAttribute( 'data-price' )));
      
        books={
          'titles':titles,
          'ids':ids,
          'prices':prices
        }
        return books;
    });
    let ti=0;
    let j=0;
    let book=[];
        for(ti=0;ti<Books.titles.length;ti++){
          
          if(match(Books.titles[ti],Title[i].BookTitle)){
            j=ti;
            book.push([Books.prices[ti],Books.ids[ti]]);
          
          }
        }
        let found;
        if(book.length===0){
          found="No";
        }
        else{
          found="yes";
        }
        if(found=="No"){
          reader.utils.sheet_add_json(worksheet, [
            { D: found, E: NaN, F: NaN,G:NaN , H:NaN, I:NaN}
          ], {skipHeader: true, origin:{r:i+1,c:3}, header: [ "D", "E", "F" ,"G","H","I"]});
          
          
          reader.utils.book_append_sheet(wb, worksheet);
         
        }
        else{
          book.sort();
        id=book[0][1];
        
        page.goto(id)
        
        let url=id;
        let price=book[0][0];
        let stock="yes";
        await page.waitForSelector('section#prodDescCont')
        let des=await page.$$eval('section#prodDescCont div#productSpecs div#id-tab-container div.tab-content  div.highlightsTileContent div.spec-body ul li ',ele=>{
          ele=ele.filter(ele =>ele.textContent.includes('Author')||ele.textContent.includes('Publisher'))
          let data=ele.map(ele=>ele.textContent)
          let author=data[1]
          let publish=data[0]
          author=author.split(":")[1];
          author=author.split("\n")[0];
          publish=publish.split(":")[1];
          publish=publish.split("\n")[0];
          let des={
            'author':author,
            'publisher':publish
          }
          return des;
        })
        let data={
          "Found":found,
          "URL":url,
          "Price":price,
          "Author":des.author,
          "Publisher":des.publisher,
          "InStock":stock
        }
        reader.utils.sheet_add_json(worksheet, [
          { D: found, E: url, F: price,G:data.Author , H:data.Publisher, I:stock}
        ], {skipHeader: true, origin:{r:i+1,c:3}, header: [ "D", "E", "F" ,"G","H","I"]});
        
        reader.utils.book_append_sheet(wb, worksheet);
        
  
        }
    }
    reader.writeFile(wb,'Input.xlsx');
    await browser.close();
    }
    catch(err){
      // console.log(err)
    }
    

      
      
  };

scrape();
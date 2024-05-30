import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.junit.jupiter.api.Test;
import org.openqa.selenium.*;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;
import org.openqa.selenium.interactions.Actions;

import java.io.File;
import java.time.Duration;
import java.time.temporal.ChronoUnit;
import java.util.*;

public class SeleniumTest {
    public static String prePath = "F:\\tianzhou\\web-data\\crawler-webofscience-npp";

    @Test
    public void hello()
    {
        System.setProperty("webdriver.edge.driver",
                "C:\\Program Files (x86)\\Microsoft\\edgeDriver\\msedgedriver.exe");
        EdgeDriver driver = new EdgeDriver();
        driver.get("http://www.baidu.com");
        driver.quit();
    }
    @Test
    public void downloadFromCNKI() throws InterruptedException {
        System.setProperty("webdriver.edge.driver",
                "C:\\Program Files (x86)\\Microsoft\\edgeDriver\\msedgedriver.exe");

        EdgeOptions edgeOptions = new EdgeOptions();
        // 允许所有请求（允许浏览器通过远程服务器访问不同源的网页，即跨域访问）
        edgeOptions.addArguments("--remote-allow-origins=*");
        EdgeDriver driver = new EdgeDriver(edgeOptions);
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(Duration.of(1, ChronoUnit.DAYS));
        List<String> keyWords1 = Arrays.asList("森林","林地","灌木","灌丛");
        List<String> keyWords2 = Arrays.asList("水","径流","产水量","水文功能","水源涵养","蒸散发","土壤水分");
        List<String> keyWords3 = Arrays.asList("净初级生产量","NPP","总初级生产量","GPP");
        for(int i=0;i< keyWords1.size();i++)
        {
            for(int j=0;j< keyWords2.size();j++)
            {
                for (int k=0;k< keyWords3.size();k++)
                {
                    downloadFromZhiwang(keyWords1.get(i), keyWords2.get(j)+";"+keyWords3.get(k),driver,"F:\\tianzhou\\web-data\\crawler-cnik-npp");
                }
            }
        }
    }
    public void downloadFromZhiwang(String forestKeyWord, String waterKeyWord, WebDriver driver, String outputPath) throws InterruptedException {
        class Paper{
            String url;
            String title;
            String author;
            String date;
            String type;

            String from;
            String abstractContent;
            String keyWords;
            String workPlace;
        }


        driver.get("https://www.cnki.net");
        WebElement element = driver.findElement(By.id("txt_SearchText"));
        element.sendKeys(forestKeyWord+";"+waterKeyWord);
        WebElement button = driver.findElement(By.className("search-btn"));
        new Actions(driver).click(button).perform();
        Thread.sleep(2000);
        driver.manage().timeouts().implicitlyWait(Duration.of(1,ChronoUnit.MINUTES));
        //获取页数
        WebElement briefBox = driver.findElement(By.id("briefBox"));
        if(briefBox.getText().contains("抱歉，暂无数据，请稍后重试。"))
        {
            System.out.println(forestKeyWord+"+"+waterKeyWord+"无数据");
            return;
        }
        WebElement paperCountElement = driver.findElement(By.className("pagerTitleCell"));
        WebElement paperCountEmElement = paperCountElement.findElement(By.tagName("em"));
        int pageCount = 1;
        try {
            List<WebElement> countPageMark = driver.findElements(By.className("countPageMark"));
            if(!countPageMark.isEmpty())
            {
                pageCount = Integer.parseInt(countPageMark.get(0).getAttribute("data-pagenum"));
            }
        }
        catch (Exception e)
        {
            e.printStackTrace();
            System.out.println("捕获到异常，没有data-pagenum元素，并设置页面为1页");
        }

        ArrayList<Paper> paperArrayList = new ArrayList<>();
        for(int i=1;i<=pageCount;i++)
        {
            try {

                WebElement resultTable = driver.findElement(By.className("result-table-list"));
                WebElement tbody = resultTable.findElement(By.tagName("tbody"));
                List<WebElement> trs = tbody.findElements(By.tagName("tr"));
                for(int j=0;j<trs.size();j++)
                {
                    //System.out.println(i+"  "+j);
                    WebElement tr = trs.get(j);
                    Paper paper = new Paper();
                    WebElement nameElement = tr.findElement(By.className("name"));
                    WebElement hrefElement = nameElement.findElement(By.tagName("a"));
                    String url = hrefElement.getAttribute("href");
                    paper.url = url;
                    WebElement dateElement = tr.findElement(By.className("date"));
                    String date = dateElement.getText();
                    paper.date = date;
                    WebElement typeElement = tr.findElement(By.className("data"));
                    WebElement typeSpanElement = typeElement.findElement(By.tagName("span"));
                    String type = typeSpanElement.getText();
                    if(!(type.equals("期刊")||type.equals("博士")||type.equals("硕士")))
                    {
                        continue;
                    }
                    paper.type = type;
                    //获取来源
                    WebElement sourceElement = tr.findElement(By.className("source"));
                    WebElement sourceAElement = sourceElement.findElement(By.tagName("a"));
                    String source = sourceAElement.getText();
                    paper.from = source;
                    paperArrayList.add(paper);
                }
                if(i<pageCount)
                {
                    WebElement nextCountButton = driver.findElement(By.id("PageNext"));
                    new Actions(driver).click(nextCountButton).perform();
                    Thread.sleep(5000);
                }
            }
            catch (Exception e)
            {
                e.printStackTrace();
                //判断是否被检测，如果被，过检测并返回true
                if(varifyCNKI(driver)){
                    System.out.println("过检测");
                }
                i--;
            }
        }
        driver.manage().timeouts().implicitlyWait(Duration.ZERO);
        for(int i=0;i<paperArrayList.size();i++)
        {
            try {
                Paper paper = paperArrayList.get(i);
                System.out.println(paper.url);
                driver.get(paper.url);
                Thread.sleep(2000);
                WebElement headElement = driver.findElement(By.className("wx-tit"));
                WebElement titleElement = headElement.findElement(By.tagName("h1"));
                String title = titleElement.getText();
                paper.title = title;
                //获取作者
                List<WebElement> titleBodyElements = headElement.findElements(By.tagName("h3"));
                WebElement authorElement = titleBodyElements.get(0);
                List<WebElement> authorsAElements = authorElement.findElements(By.tagName("a"));
                StringBuilder author = new StringBuilder();
                if(authorsAElements.size()>0)
                {
                    for(int j=0;j<authorsAElements.size();j++)
                    {
                        WebElement aElement = authorsAElements.get(j);
                        String authorTemp = aElement.getText().split("\"")[0];
                        author.append(",").append(authorTemp);
                    }
                }
                else {
                    List<WebElement> authorsSpanElements = authorElement.findElements(By.tagName("span"));
                    for(WebElement authorSpanElement:authorsSpanElements)
                    {
                        String authors = authorSpanElement.getText();
                        author.append(",").append(authors);
                    }
                }
                if(author.length()>0)
                {
                    author.deleteCharAt(0);
                }
                else {
                    author.append("无作者信息");
                }
                paper.author = author.toString();

                //获取单位
                WebElement workPlaceElement = titleBodyElements.get(1);
                List<WebElement> workPlaceAElements = workPlaceElement.findElements(By.tagName("a"));
                StringBuilder workPlace = new StringBuilder();
                if(workPlaceAElements.size()>0)
                {
                    List<WebElement> workPlaceAElement = workPlaceElement.findElements(By.tagName("a"));
                    for(int j=0;j<workPlaceAElement.size();j++)
                    {
                        WebElement aElement = workPlaceAElement.get(j);
                        String workPlaceTemp = aElement.getText();
                        workPlace.append("。").append(workPlaceTemp);
                    }

                }
                else {
                    List<WebElement> workPlacesSpanElements = workPlaceElement.findElements(By.tagName("span"));
                    for(WebElement workPlaceSpanElement:workPlacesSpanElements)
                    {
                        String workPlaceTemp = workPlaceSpanElement.getText().split("\"")[0];
                        workPlace.append("。").append(workPlaceTemp);
                    }
                }
                if(workPlace.length()>0)
                {
                    workPlace.deleteCharAt(0);
                }
                else {
                    workPlace.append("无工作单位信息");
                }
                paper.workPlace = workPlace.toString();
                //获取摘要
                WebElement abstractElement = driver.findElement(By.id("abstract_text"));
                if(!abstractElement.getAttribute("value").isBlank())
                {
                    paper.abstractContent = abstractElement.getAttribute("value");
                }
                else {
                    paper.abstractContent = "无摘要";
                }
                //获取关键词
                List<WebElement> keywordElements = driver.findElements(By.className("keywords"));
                StringBuilder keywords = new StringBuilder();
                if(keywordElements.size()>0)
                {
                    List<WebElement> keywordsAElement = keywordElements.get(0).findElements(By.tagName("a"));

                    for(int j=0;j< keywordsAElement.size();j++)
                    {
                        String keywordTemp = keywordsAElement.get(j).getText();
                        keywords.append(keywordTemp);
                    }
                }
                else {
                    keywords.append("无关键词信息");
                }
                paper.keyWords = keywords.toString();
            }
            catch (Exception e)
            {
                e.printStackTrace();
                //判断是否被检测，如果被，过检测并返回true
                if(varifyCNKI(driver)){
                    System.out.println("过检测");
                }
            }
        }
        //写入文件夹
        File directory = new File(outputPath);
        File outputFile = new File(directory,forestKeyWord+"+"+waterKeyWord+".xlsx");
        List<Map<String,String>> listRow = new ArrayList<>();
        ExcelWriter excelWriter = ExcelUtil.getWriter(outputFile);
        for (int i=0;i<paperArrayList.size();i++)
        {
            Map<String,String> map = new LinkedHashMap<>();
            Paper paper = paperArrayList.get(i);
            System.out.println(paper.title+"已完成");
            map.put("题目",paper.title);
            map.put("作者",paper.author);
            map.put("发表日期",paper.date);
            map.put("类型",paper.type);
            map.put("来源",paper.from);
            map.put("关键词",paper.keyWords);
            map.put("单位",paper.workPlace);
            map.put("摘要",paper.abstractContent);
            listRow.add(map);
        }
        excelWriter.write(listRow,true);
        excelWriter.close();
    }
    public boolean varifyCNKI(WebDriver driver) throws InterruptedException {
        try {
            if(driver.findElement(By.className("txt")).getText().contains("系统检测到您的访问行为异常"))
            {
                //获取需要滑动的距离
                WebElement varify = driver.findElement(By.className("verify-gap"));
                String left = varify.getCssValue("left");
                float move = Float.parseFloat(left.substring(0,left.length()-2));
                if(moveButton(driver,move))
                {
                    System.out.println("成功过检测");
                    return true;
                }
            }
            return false;
        }
        catch (Exception e)
        {
            System.out.println("不需要过检测");
            return true;
        }

    }
    @Test
    public boolean moveButton(WebDriver driver,float distance) throws InterruptedException {
        double x1 = -0.252689;
        double xSize = (Math.abs(x1)*2)/20;
        //distance = distance*2;
        System.out.println(distance+"");
        WebElement element = driver.findElement(By.className("verify-move-block"));
        //按下滑块
        Actions actions = new Actions(driver);
        actions.clickAndHold(element).perform();
        //往右划动
        float moved = 0;
        int i=0;
        while (Math.abs(distance-moved)>2){

            double x = x1+xSize*i;
            double y = -1*x*x+2;
            int tempMove = (int)(y*xSize*distance);
            float lastMove = moved;

            moved = moved+ tempMove;
            if(moved>distance)
            {
                actions.moveByOffset((int) (distance-lastMove),0).perform();
                moved = (int) distance;
            }
            else {
                actions.moveByOffset(tempMove,0).perform();
            }
            i++;
            System.out.println("x:"+x+"y:"+y+"index:"+i+"moved:"+moved+",,distanceTarget:"+(distance-moved));

        }
        actions.moveByOffset(10,0).perform();
        actions.release(element).perform();
        return true;
    }
    @Test
    public void downloadFromWebOfScience() throws InterruptedException {

        List<String> keywords1 = Arrays.asList("forest restoration","thinning","forest management","forestation","plantation","selective logging");
        List<String> keywords2 = Arrays.asList("water-related ecosystem services","hydrological services",
                "water supply","flow regulation","flood","peak flow","high flow","low flow","dry season flow",
                "wet season flow","water yield","streamflow","ET","soil water","soil erosion","sediment","water quality");
        //List<String> keywords3 = Arrays.asList("net primary productivity","NPP","gross primary productivity","GPP");
        int count = 0;
        ExcelReader reader = ExcelUtil.getReader(prePath+"\\task.xlsx");
        List<List<Object>> lists = reader.read();
        reader.close();
        for(int i=0;i<keywords1.size();i++)
        {
            for (int j=0;j<keywords2.size();j++)
            {
               // for(int k=0;k<keywords3.size();k++)
               // {
                    if(lists.get(count).get(2).toString().equals("1"))
                    {
                        System.out.println("已下载"+keywords1.get(i)+";"+keywords2.get(j));
                    }
                    else {
                        String[] keyword = new String[]{keywords1.get(i),keywords2.get(j)};
                        System.setProperty("webdriver.edge.driver",
                                "C:\\Program Files (x86)\\Microsoft\\edgeDriver\\msedgedriver.exe");
                        EdgeOptions edgeOptions = new EdgeOptions();
                        // 允许所有请求（允许浏览器通过远程服务器访问不同源的网页，即跨域访问）
                        edgeOptions.addArguments("--remote-allow-origins=*");
                        edgeOptions.setPageLoadStrategy(PageLoadStrategy.NORMAL);
                        EdgeDriver driver = new EdgeDriver(edgeOptions);
                        driver.manage().window().maximize();
                        HashMap<String,String> downloadMap = new HashMap<>();
                        File downloadFile = new File(prePath+"\\downloadMap\\"+keyword[0]+";"+keyword[1]+".xlsx");

                        ArrayList<Paper> papers = new ArrayList<>();
                        String output = prePath+"\\article";
                        if(downloadFile.exists())
                        {
                            ExcelReader excelReader = ExcelUtil.getReader(downloadFile);
                            List<List<Object>> read = excelReader.read();
                            for(int ii=0;ii<read.size();ii++)
                            {
                                List<Object> objects = read.get(ii);
                                downloadMap.put(objects.get(0).toString(),objects.get(1).toString());
                            }
                            excelReader.close();
                        }
                        try {
                            downloadFromWebOfScience(keyword,driver,downloadMap,papers);
                            ExcelWriter writer = ExcelUtil.getWriter(prePath+"\\task.xlsx");
                            writer.getOrCreateRow(count).getCell(2).setCellValue("1");
                            writer.close();
                        }
                        catch (Exception e){
                            e.printStackTrace();
                        }
                        finally {
                            //将downloadMap写入
                            driver.quit();
                            //写入xlsx
                            List<List<String>> lists1 = new ArrayList<>();
                            for(Map.Entry<String,String> entry:downloadMap.entrySet())
                            {
                                lists1.add(Arrays.asList(entry.getKey(),entry.getValue()));
                            }
                            ExcelWriter excelWriter = ExcelUtil.getWriter(downloadFile);
                            excelWriter.write(lists1);
                            excelWriter.close();

                            //写入文件夹
                            File directory = new File(output);
                            File outputFile = new File(directory,keyword[0]+";"+keyword[1]+".xlsx");
                            List<Map<String,String>> listRow = new ArrayList<>();
                            excelWriter = ExcelUtil.getWriter(outputFile);
                            int count2 = 0;
                            int flag =0;
                            for (int ii=0;ii<papers.size();ii++)
                            {
                                Map<String,String> map = new LinkedHashMap<>();
                                Paper paper = papers.get(ii);
                                System.out.println(paper.url+"已完成");
                                if(paper.title==null&&flag==0)
                                {
                                    count2++;
                                    continue;
                                }
                                flag=1;
                                map.put("title",paper.title);
                                map.put("author(s)",paper.author) ;
                                map.put("published",paper.date);
                                map.put("type",paper.type);
                                map.put("source",paper.from);
                                map.put("keywords",paper.keyWords);
                                map.put("addresses",paper.workPlace);
                                map.put("abstract",paper.abstractContent);
                                listRow.add(map);
                            }
                            excelWriter.passRows(count2+1);
                            boolean writeHead = true;
                            if(count2>0)
                            {
                                writeHead = false;
                            }
                            excelWriter.write(listRow,writeHead);
                            excelWriter.close();
                            //抛出异常
                            throw new RuntimeException("开启新循环");
                        }



                    }
                    count++;
               // }
            }
        }
    }
    public void downloadFromWebOfScience(String[] keyword,WebDriver driver,Map<String,String> downloadMap,ArrayList<Paper> papers) throws InterruptedException {

        driver.get("https://webofscience.clarivate.cn/wos/alldb/basic-search");
        // try {
        //     driver.findElement(By.xpath("/html/body/div[1]/div[1]/div[2]/h1/span")).getText().contains("当前无法使用此页面");
        //     System.out.println("刷新浏览器");
        //     driver.navigate().refresh();
        // }
        // catch (Exception e)
        // {
        //     throw e;
        // }
        driver.manage().timeouts().implicitlyWait(Duration.of(10,ChronoUnit.SECONDS));
        Thread.sleep(2000);
        WebElement acceptElement = driver.findElement(By.cssSelector("#onetrust-accept-btn-handler"));
        //接收cookie
        Thread.sleep(2000);
        new Actions(driver).click(acceptElement).perform();
        Thread.sleep(2000);
      //  try {
      //      Thread.sleep(2000);
      //      WebElement closeElement = driver.findElement(By.cssSelector("#pendo-close-guide-52830791"));
      //      new Actions(driver).click(closeElement).perform();
      //  }
      //  catch (Exception e)
      //  {
      //      e.printStackTrace();
      //  }

        //添加两行
        WebElement addElement = driver.findElement(By.className("add-row"));
        new Actions(driver).click(addElement).perform();
        //new Actions(driver).click(addElement).perform();
        List<WebElement> inputElements = driver.findElements(By.className("mat-form-field-infix"));
        for(int i=0;i< keyword.length;i++)
        {
            WebElement inputElement = inputElements.get(i).findElement(By.tagName("input"));
            new Actions(driver).click(inputElement).perform();
            inputElement.sendKeys("x");
        }
        List<WebElement> clearElements = driver.findElements(By.className("clear-row-button"));
        List<WebElement> selectElements = driver.findElements(By.tagName("app-select-search-field"));
        for (int i=0;i<keyword.length;i++)
        {
            new Actions(driver).click(clearElements.get(i)).perform();
            WebElement inputElement = inputElements.get(i).findElement(By.tagName("input"));
            new Actions(driver).click(inputElement).perform();
            inputElement.sendKeys(keyword[i]);
            //切换到题目
            WebElement buttonElement = selectElements.get(i).findElement(By.tagName("button"));
            new Actions(driver).click(buttonElement).perform();
            //点击Title
            WebElement input = selectElements.get(i).findElement(By.tagName("input"));
            input.sendKeys("Topic");
            input.sendKeys(Keys.ENTER);
        }
        WebElement close2Element = driver.findElement(By.id("pendo-close-guide-5600f670"));
        new Actions(driver).click(close2Element).perform();
        //找到搜索按钮
        WebElement searchElement = driver.findElement(By.cssSelector("button.search"));

        new Actions(driver).click(searchElement).perform();
        if(driver.findElements(By.cssSelector("#snSearchType > div.search-error.error-code.light-red-bg.ng-star-inserted > b")).size()>0)
        {
            System.out.println(keyword[0]+";"+keyword[1]+"无数据");
            return;
        }
        try {
            Thread.sleep(1000);
            WebElement closeElement = driver.findElement(By.cssSelector("#pendo-close-guide-30f847dd"));

            new Actions(driver).click(closeElement).perform();
        }
        catch (Exception e)
        {
            System.out.println("关闭引导界面");
            e.printStackTrace();
        }

        WebElement endPage = driver.findElement(By.cssSelector("body > app-wos > main > div > div > div.holder > div > div > div.held > app-input-route > app-base-summary-component > div > div.results.ng-star-inserted > app-page-controls.app-page-controls.ng-star-inserted > div > form > div > span"));
        int page = Integer.parseInt(endPage.getText().replaceAll(",",""));
        System.out.println("获取到"+page+"页数据");
        for(int i=1;i<=page;i++)
        {
            Thread.sleep(2000);
            WebElement resultList = driver.findElement(By.className("app-records-list"));
            List<WebElement> recordElements = resultList.findElements(By.className("app-record-holder"));
            for(int j=0;j<recordElements.size();j++)
            {
                Paper paper = new Paper();
                WebElement recordElement = recordElements.get(j);
                ((JavascriptExecutor)driver).executeScript("arguments[0].scrollIntoView();",recordElement);
                if(recordElement.getText().contains("Not available"))
                {
                    continue;
                }
                WebElement aElement = recordElement.findElement(By.cssSelector("h3>a"));
                paper.url = aElement.getAttribute("href");
                papers.add(paper);
            }
            WebElement pageElement = driver.findElement(By.className("pagination"));
            List<WebElement> buttonElements = pageElement.findElements(By.tagName("button"));
            //点击下一页
            new Actions(driver).click(buttonElements.get(1)).perform();
        }
        System.out.println("需要下载"+papers.size()+"篇文章");
        for (int i=0;i<papers.size();i++)
        {

            //System.out.println("时间为"+System.currentTimeMillis());
            Paper paper = papers.get(i);
            if(downloadMap.containsKey(paper.url))
            {
                continue;
            }


            driver.manage().timeouts().implicitlyWait(Duration.of(10,ChronoUnit.SECONDS));
            driver.get(paper.url);

            System.out.println("获取"+paper.url+"网址");
            WebElement titleElement = null;
            titleElement = driver.findElement(By.id("FullRTa-fullRecordtitle-0"));
            paper.title = titleElement.getText();
            driver.manage().timeouts().implicitlyWait(Duration.ZERO);

            //System.out.println("获取题目时间为"+System.currentTimeMillis());
            //填充作者
            //List<WebElement> authorsElement = driver.findElements(By.xpath("/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route/app-full-record-home/div[2]/div[1]/div[1]/app-full-record/div/div[2]/app-full-record-authors/div/div/span/span[1]/span[1]/span/span"));
            //                                                                             "/html/body/app-wos/main/div/div/div[2]/div/div/div[2]/app-input-route/app-full-record-home/div[2]/div[1]/div[1]/app-full-record/div/div[2]/app-full-record-authors/div/div/span/span/span/span/span";
            WebElement appFullRecord = driver.findElement(By.id("snMainArticle"));
            //WebElement authorsElement = appFullRecord.findElement(By.id(""));
            List<WebElement> authorSpanElements = appFullRecord.findElements(By.cssSelector("#SumAuthTa-MainDiv-author-en>span>span"));
            //System.out.println("一共有"+authorSpanElements.size()+"这么多作者");
            StringBuilder stringBuilder = new StringBuilder();
            for(int j=0;j<authorSpanElements.size()&&j<5;j++)
            {
                List<WebElement> authorElements = authorSpanElements.get(j).findElements(By.cssSelector("span>span>span>span>span"));
                if(authorElements.size()>0)
                {
                    stringBuilder.append((j+1)+"-").append(authorElements.get(0).getText()).append("\n");
                }
                //韩国人类型
                else {
                    WebElement authorElement = authorSpanElements.get(j).findElement(By.cssSelector(".authors.value"));
                    stringBuilder.append((j+1)+"-").append(authorElement.getText()).append("\n");
                }

            }
            paper.author = stringBuilder.toString();
            //System.out.println("获取作者时间为"+System.currentTimeMillis());
            //填充来源
            List<WebElement> divElements = appFullRecord.findElements(By.tagName("div"));
            for(WebElement div:divElements)
            {
                if(div.getText().contains("Source")){
                    List<WebElement> sourceElements = div.findElements(By.tagName("mat-sidenav-content"));
                    if(sourceElements.size()>0)
                    {
                        paper.from = sourceElements.get(0).getText();
                        break;
                    }
                }
            }
            //System.out.println("获取来源时间为"+System.currentTimeMillis());
            //填充发表日期
            List<WebElement> publishElement = appFullRecord.findElements(By.cssSelector("#FullRTa-pubdate"));
            if(publishElement.size()>0)
            {
                paper.date = publishElement.get(0).getText();
            }

            //System.out.println("获取发表时间为"+System.currentTimeMillis());
            //填充类型
            WebElement typeElement = appFullRecord.findElement(By.cssSelector("#FullRTa-doctype-0"));
            paper.type = typeElement.getText();
            //System.out.println("获取类型时间为"+System.currentTimeMillis());
            //填充关键词
            List<WebElement> elements = appFullRecord.findElements(By.cssSelector("app-full-record-keywords"));
            //System.out.println("获取关键词1时间为"+System.currentTimeMillis());
            List<WebElement> contentElement = elements.get(0).findElements(By.cssSelector("div"));
            //System.out.println("获取关键词2时间为"+System.currentTimeMillis());
            List<WebElement> authorElementH2 = appFullRecord.findElements(By.cssSelector("div>#FullRTa-pqdt_authorKeyWords"));
            if(contentElement.size()>0)
            {
                List<WebElement> keywordDivElements = elements.get(0).findElements(By.cssSelector("span>div"));

                List<WebElement> spanElements = keywordDivElements.get(0).findElements(By.cssSelector("span>a>span"));

                stringBuilder = new StringBuilder();
                for(int j=0;j< spanElements.size();j++)
                {
                    stringBuilder.append((j+1)+"-").append(spanElements.get(j).getText()).append("\n");
                }
                paper.keyWords = stringBuilder.toString();
            }
            else if(authorElementH2.size()>0)
            {
                for(WebElement div:divElements)
                {
                    if(div.getText().contains("Author Keywords")){
                        List<WebElement> valueElements = div.findElements(By.cssSelector("div>span>.value"));
                        for(int j=0;j< valueElements.size();j++)
                        {
                            stringBuilder.append((j+1)+"-").append(valueElements.get(j).getText()).append("\n");
                        }
                        paper.keyWords = stringBuilder.toString();
                        break;
                    }
                    //  if(!div.getAttribute("class").equals("cdx-two-column-grid-container ng-star-inserted"))
                    //  {
                    //      continue;
                    //  }
                    //  List<WebElement> authorKeyWordsElements = div.findElements(By.cssSelector("div>#FullRTa-pqdt_authorKeyWords"));
                    //  if(authorKeyWordsElements.size()>0)
                    //  {
                    //      stringBuilder = new StringBuilder();
                    //      List<WebElement> spanElements = div.findElements(By.cssSelector("div>span>.value"));
                    //      for(int j=0;j< spanElements.size();j++)
                    //      {
                    //          stringBuilder.append((j+1)+"-").append(spanElements.get(j).getText()).append("\n");
                    //      }
                    //      paper.keyWords = stringBuilder.toString();
                    //      break;
                    //  }
                }
            }
            //System.out.println("获取关键词时间为"+System.currentTimeMillis());
            // else {
            //     WebElement spanElement = appFullRecord.findElement(By.xpath("div[9]/span"));
            //     List<WebElement> spanElements = spanElement.findElements(By.cssSelector("span.value.ng-star-inserted"));
            //     stringBuilder = new StringBuilder();
            //     for(int j=0;j< spanElements.size();j++)
            //     {
            //         stringBuilder.append((j+1)+"-").append(spanElements.get(j).getText()).append("\n");
            //     }
            //     paper.keyWords = stringBuilder.toString();
            // }

            // try {
            //     WebElement keywordsElement = driver.findElement(By.cssSelector("#snMainArticle > app-full-record-keywords > div > span > div.ng-star-inserted"));
            //     List<WebElement> aKeywordElements = keywordsElement.findElements(By.tagName("a"));
            //     stringBuilder = new StringBuilder();
            //     for(int j=0;j<aKeywordElements.size();j++)
            //     {
            //         stringBuilder.append("|").append(aKeywordElements.get(j).getText());
            //     }
            //     if(stringBuilder.length()>0)
            //     {
            //         stringBuilder.deleteCharAt(0);
            //     }
            //     paper.keyWords = stringBuilder.toString();
            // }
            // catch (Exception e)
            // {
            //     WebElement keywordsElement = driver.findElement(By.xpath("app-full-record"));
            //     List<WebElement> keywordSpanElements = keywordsElement.findElements(By.tagName("span"));
            //     stringBuilder = new StringBuilder();
            //     for(int j=0;j<keywordSpanElements.size();j++)
            //     {
            //         stringBuilder.append("|").append(keywordSpanElements.get(j).getText());
            //     }
            //     if(stringBuilder.length()>0)
            //     {
            //         stringBuilder.deleteCharAt(0);
            //     }
            //     paper.keyWords = stringBuilder.toString();
            // }


            //填充单位
            stringBuilder = new StringBuilder();
            List<WebElement> addressElements = driver.findElements(By.cssSelector("app-full-record-author-organization span.value.padding-right-5--reversible"));
            for(int j=0;j<addressElements.size()&&j<5;j++)
            {
                stringBuilder.append((j+1)+"-").append(addressElements.get(j).getText()).append("\n");
            }
            paper.workPlace = stringBuilder.toString();
            //System.out.println("获取单位时间为"+System.currentTimeMillis());
            //填充摘要
            List<WebElement> abstractElement = appFullRecord.findElements(By.id("FullRTa-abstract-basic"));
            if(abstractElement.size()>0)
            {
                paper.abstractContent = abstractElement.get(0).getText();
            }
            downloadMap.put(paper.url,"1");
            //System.out.println("获取摘要时间为"+System.currentTimeMillis());
            // WebElement abstractElement = driver.findElement(By.cssSelector("#snMainArticle > div:nth-child(9) > div > div"));
            // paper.abstractContent = abstractElement.getText();
        }


    }
    @Test
    public void testSignal() throws InterruptedException {
        String keyword = "Mangroves and Sediment Dynamics Along the Coasts of Southern Thailand";
        System.setProperty("webdriver.edge.driver",
                "C:\\Program Files (x86)\\Microsoft\\edgeDriver\\msedgedriver.exe");
        EdgeOptions edgeOptions = new EdgeOptions();
        // 允许所有请求（允许浏览器通过远程服务器访问不同源的网页，即跨域访问）
        edgeOptions.addArguments("--remote-allow-origins=*");
        edgeOptions.setPageLoadStrategy(PageLoadStrategy.NORMAL);
        EdgeDriver driver = new EdgeDriver(edgeOptions);
        driver.manage().window().maximize();
        downloadFromWebOfScience(new String[]{keyword,""},driver,new HashMap<>(),new ArrayList<Paper>());
        driver.quit();
    }
    @Test
    public void loop() throws InterruptedException {
        ZipSecureFile.setMinInflateRatio(-1);
        while (true)
        {
            try {
                if(downloadOver()==false)
                {
                    downloadFromWebOfScience();
                }
                else {
                    break;
                }
            }
            catch (Exception e)
            {
                e.printStackTrace();
                System.out.println("进入下一次循环");
            }
        }
    }
    public boolean downloadOver()
    {
        ExcelReader reader = ExcelUtil.getReader(prePath+"\\task.xlsx");
        List<List<Object>> lists = reader.read();
        for(int i=0;i<lists.size();i++)
        {
            if(lists.get(i).get(2).toString().equals("0"))
            {
                reader.close();
                return false;
            }
        }
        reader.close();
        return true;

    }
    @Test
    public void makeTaskExcel()
    {
        ExcelWriter writer = ExcelUtil.getWriter(prePath+"\\task.xlsx");
        List<List<String>> lists = new ArrayList<>();
        List<String> keywords1 = Arrays.asList("forest restoration","thinning","forest management","forestation","plantation","selective logging");
        List<String> keywords2 = Arrays.asList("water-related ecosystem services","hydrological services",
                "water supply","flow regulation","flood","peak flow","high flow","low flow","dry season flow",
                "wet season flow","water yield","streamflow","ET","soil water","soil erosion","sediment","water quality");
        //List<String> keywords3 = Arrays.asList("net primary productivity","NPP","gross primary productivity","GPP");
        for(int i=0;i<keywords1.size();i++)
        {
            for (int j=0;j<keywords2.size();j++)
            {
               // for(int k=0;k<keywords3.size();k++)
               // {
                    lists.add(Arrays.asList(keywords1.get(i),keywords2.get(j),"0"));
               // }
            }
        }
        writer.write(lists);
        writer.close();
    }
    public void downloadFromJournal(String keyword,String journal,EdgeDriver driver,String outputPath) throws InterruptedException {
        driver.manage().timeouts().implicitlyWait(Duration.of(10,ChronoUnit.SECONDS));
        driver.get("https://kns.cnki.net/kns8s/AdvSearch");
        WebElement dlElement = driver.findElement(By.id("gradetxt"));
        List<WebElement> ddElements = dlElement.findElements(By.tagName("dd"));
        //第一个是主题
        WebElement topicElement = ddElements.get(0);
        WebElement inputTopicElement = topicElement.findElement(By.cssSelector(".input-box>input"));
        inputTopicElement.sendKeys(keyword);
        Thread.sleep(1000);
        //第二个是来源
        WebElement sourceElement = ddElements.get(1);
        WebElement chooseElement = sourceElement.findElements(By.className("sort-default")).get(1);
        new Actions(driver).click(chooseElement).perform();
        Thread.sleep(1000);
            //选择参考文献
        List<WebElement> liElements = sourceElement.findElements(By.cssSelector(".reopt>.sort-list>ul>li"));
        new Actions(driver).click(liElements.get(14)).perform();
        WebElement inputSourceElement = sourceElement.findElement(By.cssSelector(".input-box>input"));
        inputSourceElement.sendKeys(journal);
        Thread.sleep(1000);
        //点击确定
        WebElement search = driver.findElement(By.cssSelector("input.btn-search"));
        new Actions(driver).click(search).perform();
        Thread.sleep(5000);
        //获取页数
        WebElement briefBox = driver.findElement(By.id("briefBox"));
        if(briefBox.getText().contains("抱歉，暂无数据，请稍后重试。"))
        {
            System.out.println(keyword+"无数据");
            return;
        }
        WebElement paperCountElement = driver.findElement(By.className("pagerTitleCell"));
        WebElement paperCountEmElement = paperCountElement.findElement(By.tagName("em"));
        int pageCount = 1;
        try {
            List<WebElement> countPageMark = driver.findElements(By.className("countPageMark"));
            if(!countPageMark.isEmpty())
            {
                pageCount = Integer.parseInt(countPageMark.get(0).getAttribute("data-pagenum"));
            }
        }
        catch (Exception e)
        {
            e.printStackTrace();
            System.out.println("捕获到异常，没有data-pagenum元素，并设置页面为1页");
        }

        ArrayList<Paper> paperArrayList = new ArrayList<>();
        for(int i=1;i<=pageCount;i++)
        {
            try {

                WebElement resultTable = driver.findElement(By.className("result-table-list"));
                WebElement tbody = resultTable.findElement(By.tagName("tbody"));
                List<WebElement> trs = tbody.findElements(By.tagName("tr"));
                for(int j=0;j<trs.size();j++)
                {
                    //System.out.println(i+"  "+j);
                    WebElement tr = trs.get(j);
                    Paper paper = new Paper();
                    WebElement nameElement = tr.findElement(By.className("name"));
                    WebElement hrefElement = nameElement.findElement(By.tagName("a"));
                    String url = hrefElement.getAttribute("href");
                    paper.url = url;
                    WebElement dateElement = tr.findElement(By.className("date"));
                    String date = dateElement.getText();
                    paper.date = date;
                    WebElement typeElement = tr.findElement(By.className("data"));
                    WebElement typeSpanElement = typeElement.findElement(By.tagName("span"));
                    String type = typeSpanElement.getText();
                    if(!(type.equals("期刊")||type.equals("博士")||type.equals("硕士")))
                    {
                        continue;
                    }
                    paper.type = type;
                    //获取来源
                    WebElement sourceElement2 = tr.findElement(By.className("source"));
                    WebElement sourceAElement = sourceElement2.findElement(By.tagName("a"));
                    String source = sourceAElement.getText();
                    paper.from = source;
                    paperArrayList.add(paper);
                }
                if(i<pageCount)
                {
                    WebElement nextCountButton = driver.findElement(By.id("PageNext"));
                    new Actions(driver).click(nextCountButton).perform();
                    Thread.sleep(5000);
                }
            }
            catch (Exception e)
            {
                e.printStackTrace();
                //判断是否被检测，如果被，过检测并返回true
                if(varifyCNKI(driver)){
                    System.out.println("过检测");
                }
                i--;
            }
        }
        driver.manage().timeouts().implicitlyWait(Duration.ZERO);
        for(int i=0;i<paperArrayList.size();i++)
        {
            try {
                Paper paper = paperArrayList.get(i);
                System.out.println(paper.url);
                Thread.sleep(1000);
                driver.get(paper.url);
                Thread.sleep(1000);
                WebElement headElement = driver.findElement(By.className("wx-tit"));
                WebElement titleElement = headElement.findElement(By.tagName("h1"));
                String title = titleElement.getText();
                paper.title = title;
                //获取作者
                List<WebElement> titleBodyElements = headElement.findElements(By.tagName("h3"));
                WebElement authorElement = titleBodyElements.get(0);
                List<WebElement> authorsAElements = authorElement.findElements(By.tagName("a"));
                StringBuilder author = new StringBuilder();
                if(authorsAElements.size()>0)
                {
                    for(int j=0;j<authorsAElements.size();j++)
                    {
                        WebElement aElement = authorsAElements.get(j);
                        String authorTemp = aElement.getText().split("\"")[0];
                        author.append(",").append(authorTemp);
                    }
                }
                else {
                    List<WebElement> authorsSpanElements = authorElement.findElements(By.tagName("span"));
                    for(WebElement authorSpanElement:authorsSpanElements)
                    {
                        String authors = authorSpanElement.getText();
                        author.append(",").append(authors);
                    }
                }
                if(author.length()>0)
                {
                    author.deleteCharAt(0);
                }
                else {
                    author.append("无作者信息");
                }
                paper.author = author.toString();

                //获取单位
                WebElement workPlaceElement = titleBodyElements.get(1);
                List<WebElement> workPlaceAElements = workPlaceElement.findElements(By.tagName("a"));
                StringBuilder workPlace = new StringBuilder();
                if(workPlaceAElements.size()>0)
                {
                    List<WebElement> workPlaceAElement = workPlaceElement.findElements(By.tagName("a"));
                    for(int j=0;j<workPlaceAElement.size();j++)
                    {
                        WebElement aElement = workPlaceAElement.get(j);
                        String workPlaceTemp = aElement.getText();
                        workPlace.append("。").append(workPlaceTemp);
                    }

                }
                else {
                    List<WebElement> workPlacesSpanElements = workPlaceElement.findElements(By.tagName("span"));
                    for(WebElement workPlaceSpanElement:workPlacesSpanElements)
                    {
                        String workPlaceTemp = workPlaceSpanElement.getText().split("\"")[0];
                        workPlace.append("。").append(workPlaceTemp);
                    }
                }
                if(workPlace.length()>0)
                {
                    workPlace.deleteCharAt(0);
                }
                else {
                    workPlace.append("无工作单位信息");
                }
                paper.workPlace = workPlace.toString();
                //获取摘要
                WebElement abstractElement = driver.findElement(By.id("abstract_text"));
                if(!abstractElement.getAttribute("value").isBlank())
                {
                    paper.abstractContent = abstractElement.getAttribute("value");
                }
                else {
                    paper.abstractContent = "无摘要";
                }
                //获取关键词
                List<WebElement> keywordElements = driver.findElements(By.className("keywords"));
                StringBuilder keywords = new StringBuilder();
                if(keywordElements.size()>0)
                {
                    List<WebElement> keywordsAElement = keywordElements.get(0).findElements(By.tagName("a"));

                    for(int j=0;j< keywordsAElement.size();j++)
                    {
                        String keywordTemp = keywordsAElement.get(j).getText();
                        keywords.append(keywordTemp);
                    }
                }
                else {
                    keywords.append("无关键词信息");
                }
                paper.keyWords = keywords.toString();
            }
            catch (Exception e)
            {
                e.printStackTrace();
                //判断是否被检测，如果被，过检测并返回true
                if(varifyCNKI(driver)){
                    System.out.println("过检测");
                }
            }
        }
        driver.close();
        //写入文件夹
        File directory = new File(outputPath);
        File outputFile = new File(directory,keyword+".xlsx");
        List<Map<String,String>> listRow = new ArrayList<>();
        ExcelWriter excelWriter = ExcelUtil.getWriter(outputFile);
        for (int i=0;i<paperArrayList.size();i++)
        {
            Map<String,String> map = new LinkedHashMap<>();
            Paper paper = paperArrayList.get(i);
            System.out.println(paper.title+"已完成");
            map.put("题目",paper.title);
            map.put("作者",paper.author);
            map.put("发表日期",paper.date);
            map.put("类型",paper.type);
            map.put("来源",paper.from);
            map.put("关键词",paper.keyWords);
            map.put("单位",paper.workPlace);
            map.put("摘要",paper.abstractContent);
            listRow.add(map);
        }
        excelWriter.write(listRow,true);
        excelWriter.close();
        //
        //写入excel
    }
    @Test
    public void downloadFromEcologyJournal() throws InterruptedException {
        String journalName = "生态学报";
        List<String> keywords = Arrays.asList("森林固碳量","森林碳储量","森林碳汇","四川森林碳储量","四川森林固碳量","四川森林碳汇","西南地区森林碳储量","西南地区森林固碳量","西南地区碳汇");
        String outputPath = "F:\\tianzhou\\web-data\\crawler-生态学报";
        for(int i=0;i<keywords.size();i++)
        {
            System.setProperty("webdriver.edge.driver",
                    "C:\\Program Files (x86)\\Microsoft\\edgeDriver\\msedgedriver.exe");
            EdgeOptions edgeOptions = new EdgeOptions();
            // 允许所有请求（允许浏览器通过远程服务器访问不同源的网页，即跨域访问）
            edgeOptions.addArguments("--remote-allow-origins=*");
            //edgeOptions.setPageLoadStrategy(PageLoadStrategy.NORMAL);
            EdgeDriver driver = new EdgeDriver(edgeOptions);
            driver.manage().window().maximize();
            downloadFromJournal(keywords.get(i),journalName,driver,outputPath);
        }
    }
    @Test
    public void testExcelWrite()
    {
        String filePath = "F:\\tianzhou\\web-data\\crawler-webofscience-npp\\test.xlsx";
        ExcelWriter excelWriter = ExcelUtil.getWriter(filePath);
        List<Map> list = new ArrayList<>();
        Map<String,String> map = new HashMap<>();
        map.put("x","x");
        list.add(map);
        excelWriter.passRows(4);
        excelWriter.write(list,false);
        excelWriter.close();
    }
}
class Paper{
    String url;
    String title;
    String author;
    String date;
    String type;

    String from;
    String abstractContent;
    String keyWords;
    String workPlace;
}

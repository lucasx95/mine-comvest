import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.stream.IntStream;


/**
 * Classe para minerar dados sócioeconômicos dos matriculados na unicamp
 * de 1987 a 2016
 */
public class MineComvest {

    private static WebDriver driver;
    private static Workbook wb;
    private static CreationHelper createHelper;


    /**
     * Finaliza o processo de mineração dos dados
     *
     * @param title o título que o arquivo vai receber
     * @throws IOException
     */
    private static void closeMining(String title) throws IOException {

        /**
         * Cria arquivo de saída e grava em disco
         */
        FileOutputStream fileOut = new FileOutputStream(title);
        wb.write(fileOut);
        fileOut.close();

        // Fecha o driver
        driver.close();
    }

    /**
     * Preenche as linhas das planilhas
     *
     * @param sheet        a planilha para adicionar a linha
     * @param headerFields os campos do cabeçalho
     */
    private static void fillRow(Sheet sheet, ArrayList<String> headerFields, short rowNumber) {


        // Cria estilo das células do cabeçalho
        final CellStyle style = wb.createCellStyle();
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);

        // Decide a cor da linha de acordo se é cabeçalho, impar ou par
        if (rowNumber == 0) {
            style.setFillForegroundColor(HSSFColor.LEMON_CHIFFON.index);
        } else if (rowNumber % 2 == 0) {
            style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        }

        // Cria a linha para ser inserida
        final Row row = sheet.createRow(rowNumber);

        // Preenche a linha
        IntStream.range(0, headerFields.size()).forEach(i -> {
            row.createCell(i).setCellValue(headerFields.get(i));
            row.getCell(i).setCellStyle(style);
        });

    }

    /**
     * Método para coletar porcentagens dos dados desejados
     *
     * @param percentageType texto que descreve porcentagem a ser coletada
     * @return a porcentagem, ou "-" caso o tipo seja vazio
     */
    private static String collectPercentage(String percentageType) {
        if ("".equals(percentageType)) return "-";
        return driver.findElement(By.xpath("//*[text()[contains(.,'" + percentageType + "')]]/following-sibling::td" +
                "[@class='tabelatexto']/following-sibling::td[@class='tabelatexto']")).getText();
    }


    /**
     * Minera os dados relativos ao tipo de ensino médio cursado
     */
    private static void mineHighSchool() throws IOException {

        //Contador para a linha sendo preenchida
        short row = 0;

        // Cria tabela
        final Sheet sheet = wb.createSheet("Matriculados Unicamp - Ensino Médio - 1987 à 2016");

        // Cria lista de dados a serem coletados e inicia com o cabeçalho
        final ArrayList<String> rowData = new ArrayList<>(Arrays.asList("Ano", "Em Branco", "Particular", "Público",
                "Maioria Público", "Maioria Particular", "Dividido", "Exterior", "N.D.A"));

        // Preenche o cabeçalho
        fillRow(sheet, rowData, row++);

        // Inicializa Strings que identifica a questão e os dados
        String question = "Em que tipo de estabelecimento de ensino você cursou o 2º grau";
        String blankAnswer = "em branco";
        String exclusivePrivate = "cursei somente em estabelecimento particular";
        String exclusivePublic = "cursei somente em estabelecimento público";
        String mixPublic = "cursei parte em esc. pública e parte em esc. particular, ficando mais em esc. pública";
        String mixPrivate = "cursei parte em esc. particular e parte em esc. pública, ficando mais em esc. particular";
        String mixDivided = "";
        String otherCountry = "";
        String nonAlternative = "nenhuma das alternativas anteriores";

        // Coleta dados de todos os anos
        for (int i = 1987; i < 2017; i++) {

            // Em alguns anos os identificadores mudam e são alterados
            if (1989 == i) {
                mixDivided = "cursei parte em esc. particular e parte em esc. pública, ficando igual intervalo de tempo";
            } else if (1999 == i) {
                question = "Em que tipo de estabelecimento você cursou o 2º grau";
            } else if (2000 == i) {
                question = "Em que tipo de estabelecimento você cursou o ensino médio";
            } else if (2005 == i) {
                question = "Em que estabelecimento você cursou o ensino médio";
                exclusivePrivate = "somente particular";
                exclusivePublic = "somente público";
                mixPublic = "misto, mais tempo em estabelecimento público";
                mixPrivate = "misto, mais tempo em estabelecimento particular";
                mixDivided = "misto, em igual intervalo de tempo";
            } else if (2013 == i) {
                question = "Onde você cursou o ensino médio?";
                exclusivePrivate = "todo em escola particular";
                exclusivePublic = "todo em escola pública";
                mixPublic = "maior parte em escola pública";
                mixPrivate = "maior parte em escola particular";
                mixDivided = "";
                otherCountry = "no exterior";
                nonAlternative = "em outra situação";
            }

            // Avança para a tela de dados de acordo com o ano
            driver.navigate().to("https://www.comvest.unicamp.br/estatisticas/" + String.valueOf(i) + "/quest/quest1.php");
            final Select questionBox = new Select(driver.findElement(By.name("questao")));
            driver.findElement(By.xpath("//*[text()[contains(.,'" + question + "')]]\n")).click();
            driver.findElement(By.xpath("//*[text()[contains(.,'Matriculados')]]\n")).click();
            driver.findElement(By.name("Executar")).click();

            // Coleta os dados
            rowData.set(0, String.valueOf(i));
            rowData.set(1, collectPercentage(blankAnswer));
            rowData.set(2, collectPercentage(exclusivePrivate));
            rowData.set(3, collectPercentage(exclusivePublic));
            rowData.set(4, collectPercentage(mixPublic));
            rowData.set(5, collectPercentage(mixPrivate));
            rowData.set(6, collectPercentage(mixDivided));
            rowData.set(7, collectPercentage(otherCountry));
            rowData.set(8, collectPercentage(nonAlternative));

            // Preenche os dados
            fillRow(sheet, rowData, row++);
        }

        // Seta o tamanho de todas as colunas para sua maior célula
        IntStream.range(0, rowData.size()).forEach(sheet::autoSizeColumn);

    }

    /**
     * Minera dados relativos a Raça ou Cor
     *
     * @throws IOException
     */
    private static void mineRaceOrColor() throws IOException {
        //Contador para a linha sendo preenchida
        short row = 0;

        // Cria tabela
        final Sheet sheet = wb.createSheet("Matriculados Unicamp - Raça ou Cor - 2003 à 2016");

        // Cria lista de dados a serem coletados e inicia com o cabeçalho
        final ArrayList<String> rowData = new ArrayList<>(Arrays.asList(
                "Ano", "Em Branco", "Branca", "Preta", "Parda", "Amarela", "Indígena", "Não Declarada"));

        // Preenche o cabeçalho
        fillRow(sheet, rowData, row++);

        // Inicializa Strings que identifica a questão e os dados
        String question = "cor ou raça";
        String blankAnswer = "em branco";
        String blank = "branca";
        String black = "preta";
        String brown = "parda";
        String yellow = "amarela";
        String indian = "indígena";
        String undeclared = "";

        for (int i = 2003; i < 2017; i++) {
            // Em 2013 foi adicionada a opção não declarada
            if (2013 == i) undeclared = "não declarada";

            // Avança para a tela de dados de acordo com o ano
            driver.navigate().to("https://www.comvest.unicamp.br/estatisticas/" + String.valueOf(i) + "/quest/quest1.php");
            final Select questionBox = new Select(driver.findElement(By.name("questao")));
            driver.findElement(By.xpath("//*[text()[contains(.,'" + question + "')]]\n")).click();
            driver.findElement(By.xpath("//*[text()[contains(.,'Matriculados')]]\n")).click();
            driver.findElement(By.name("Executar")).click();

            // Coleta os dados
            rowData.set(0, String.valueOf(i));
            rowData.set(1, collectPercentage(blankAnswer));
            rowData.set(2, collectPercentage(blank));
            rowData.set(3, collectPercentage(black));
            rowData.set(4, collectPercentage(brown));
            rowData.set(5, collectPercentage(yellow));
            rowData.set(6, collectPercentage(indian));
            rowData.set(7, collectPercentage(undeclared));

            // Preenche os dados
            fillRow(sheet, rowData, row++);
        }

        // Seta o tamanho de todas as colunas para sua maior célula
        IntStream.range(0, rowData.size()).forEach(sheet::autoSizeColumn);

    }

    /**
     * Minera dados relativos a renda Mensal Total recebida
     */
    private static void mineMonthlyIncome() throws IOException {
        //Contador para a linha sendo preenchida
        short row = 0;

        // Cria tabela
        final Sheet sheet = wb.createSheet("Matriculados Unicamp - Renda Mensal Total em Salários Mínimos - 2013 à 2016");

        // Cria lista de dados a serem coletados e inicia com o cabeçalho
        final ArrayList<String> rowData = new ArrayList<>(Arrays.asList(
                "Ano", "Em Branco", "Até um 1 S.M ", "Entre 1 e 2 S.M", "Entre 2 e 3 S.M"
                , "Entre 3 e 5 S.M", "Entre 5 e 7 S.M", "Entre 7 e 10 S.M", "Entre 10 e 15 S.M", "Entre 15 e 20 S.M", "Acima de 20 S.M"));

        // Preenche o cabeçalho
        fillRow(sheet, rowData, row++);

        // Inicializa Strings que identifica a questão e os dados
        final String question = "renda mensal total";
        final String blankAnswer = "em branco";
        final String lessOne = "inferior a 01 SM";
        final String betweenOneTwo = "entre 01 e 02 SM";
        final String betweenTwoThree = "entre 02 e 03 SM";
        final String betweenThreeFive = "entre 03 e 05 SM";
        final String betweenFiveSeven = "entre 05 e 07 SM";
        final String betweenSevenTen = "entre 07 e 10 SM";
        final String betweenTenFifteen = "entre 10 e 15 SM";
        final String betweenFifteenTwenty = "entre 15 e 20 SM";
        final String biggerTwenty = "acima de 20 SM";


        for (int i = 2013; i < 2017; i++) {

            // Avança para a tela de dados de acordo com o ano
            driver.navigate().to("https://www.comvest.unicamp.br/estatisticas/" + String.valueOf(i) + "/quest/quest1.php");
            final Select questionBox = new Select(driver.findElement(By.name("questao")));
            driver.findElement(By.xpath("//*[text()[contains(.,'" + question + "')]]\n")).click();
            driver.findElement(By.xpath("//*[text()[contains(.,'Matriculados')]]\n")).click();
            driver.findElement(By.name("Executar")).click();

            // Coleta os dados
            rowData.set(0, String.valueOf(i));
            rowData.set(1, collectPercentage(blankAnswer));
            rowData.set(2, collectPercentage(lessOne));
            rowData.set(3, collectPercentage(betweenOneTwo));
            rowData.set(4, collectPercentage(betweenTwoThree));
            rowData.set(5, collectPercentage(betweenThreeFive));
            rowData.set(6, collectPercentage(betweenFiveSeven));
            rowData.set(7, collectPercentage(betweenSevenTen));
            rowData.set(8, collectPercentage(betweenTenFifteen));
            rowData.set(9, collectPercentage(betweenFifteenTwenty));
            rowData.set(10, collectPercentage(biggerTwenty));

            // Preenche os dados
            fillRow(sheet, rowData, row++);
        }

        sheet.createRow(row).createCell(0).setCellValue(" Somente de 2013 para frente, pois anteriormente era utilizada outra escala.");

        // Seta o tamanho de todas as colunas para sua maior célula
        IntStream.range(0, rowData.size()).forEach(sheet::autoSizeColumn);

    }

    /**
     * Minera dados relativos ao número de livros na casa além dos didáticos
     *
     * @throws IOException
     */
    private static void mineBooksNumber() throws IOException {
        //Contador para a linha sendo preenchida
        short row = 0;

        // Cria tabela
        final Sheet sheet = wb.createSheet("Matriculados Unicamp -  Número de livros na casa além dos didáticos - 2013 à 2016");

        // Cria lista de dados a serem coletados e inicia com o cabeçalho
        final ArrayList<String> rowData = new ArrayList<>(Arrays.asList(
                "Ano", "Em Branco", "Nenhum", "1 a 20 livros", "21 a 100 livros", "Mais de 100 livros"));

        // Preenche o cabeçalho
        fillRow(sheet, rowData, row++);

        // Inicializa Strings que identifica a questão e os dados
        final String question = "dos livros escolares";
        final String blankAnswer = "em branco";
        final String none = "nenhum";
        final String betweenOneTwenty = "1 a 20 livros";
        final String betweenTwentyOneAHundred = "21 a 100 livros";
        final String biggerHundred = "mais de 100 livros";

        for (int i = 2004; i < 2017; i++) {

            // Avança para a tela de dados de acordo com o ano
            driver.navigate().to("https://www.comvest.unicamp.br/estatisticas/" + String.valueOf(i) + "/quest/quest1.php");
            final Select questionBox = new Select(driver.findElement(By.name("questao")));
            driver.findElement(By.xpath("//*[text()[contains(.,'" + question + "')]]\n")).click();
            driver.findElement(By.xpath("//*[text()[contains(.,'Matriculados')]]\n")).click();
            driver.findElement(By.name("Executar")).click();

            // Coleta os dados
            rowData.set(0, String.valueOf(i));
            rowData.set(1, collectPercentage(blankAnswer));
            rowData.set(2, collectPercentage(none));
            rowData.set(3, collectPercentage(betweenOneTwenty));
            rowData.set(4, collectPercentage(betweenTwentyOneAHundred));
            rowData.set(5, collectPercentage(biggerHundred));

            // Preenche os dados
            fillRow(sheet, rowData, row++);
        }

        // Seta o tamanho de todas as colunas para sua maior célula
        IntStream.range(0, rowData.size()).forEach(sheet::autoSizeColumn);

    }

    public static void main(String[] args) throws InterruptedException, IOException {

        // Inicializa o Webdriver
        driver = new FirefoxDriver();

        //Inicializa variáveis para persistir os dados
        wb = new HSSFWorkbook();
        createHelper = wb.getCreationHelper();

        // Minera os dados
        mineRaceOrColor();
        mineHighSchool();
        mineMonthlyIncome();
        mineBooksNumber();

        // Finaliza processo de coleta de dados
        closeMining("Dados_ComvestUnicamp.xls");
    }
}

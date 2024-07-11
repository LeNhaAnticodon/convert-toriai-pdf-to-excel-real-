package com.example.convert_toriai_pdf_to_excel.convert;

import com.example.convert_toriai_pdf_to_excel.model.CsvFile;
import com.opencsv.CSVWriter;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.TimeoutException;

public class ReadPDFToExcel {

    private static final Set<Double> seiHinSet = new LinkedHashSet<>();
    // list chứa danh sách các sản phẩm không trùng lặp
    ObservableList<Double> seiHinList = FXCollections.observableArrayList(seiHinSet);
    // time tháng và ngày
    private static String shortNouKi = "";
    // 備考
    private static String bikou = "";
    // 客先名
    private static String kyakuSakiMei = "";
    // 3 kích thước của vật liệu
    private static int size1;
    private static int size2;
    private static int size3 = 0;
    // ký hiệu loại vật lệu
    private static String koSyuNumMark = "3";
    // 切りロス
    private static String kirirosu = "";

    // tên file chl sẽ tạo được ghi trong phần 工事名, chưa bao gồm loại vật liệu
    public static String fileExcelName = "name";

    // link của file pdf
    private static String pdfPath = "";

    // link thư mục của file excel xlsx sẽ tạo
    private static String xlsxExcelPath = "";
    // link thư mục của file excel csv sẽ tạo
    private static String csvExcelDirPath = "";
    // link thư mục của file chl sẽ tạo
    private static String chlDirPath = "";
    // đếm số dòng sẽ tạo trên file chl
    private static int rowToriAiNum;

    // loại vật liệu và kích thước
    private static String kouSyu;

    // tên loại vật liệu
    private static String kouSyuName;
    // tên file chl đầy đủ sẽ tạo đã bao gồm tên loại vật liệu
    public static String fileName;
    private static String excelPath;
    private static String chuyuBan = "";
    private static String teiHaiSha = "";
    private static String tanZyuu;
    private static List<Double> listBoZai = new ArrayList<>();
    private static List<Double> listSeiHin = new ArrayList<>();

    /**
     * chuyển đổi pdf tính vật liệu thành các file chl theo từng vật liệu khác nhau
     *
     * @param filePDFPath    link file pdf
     * @param fileChlDirPath link thư mục chứa file chl sẽ tạo
     * @param csvFileNames   list chứa danh sách các file chl đã tạo
     */
    public static void convertPDFToExcel(String filePDFPath, String fileChlDirPath, ObservableList<CsvFile> csvFileNames) throws FileNotFoundException, TimeoutException, IOException {
        // xóa danh sách cũ trước khi thực hiện, tránh bị ghi chồng lên nhau
        csvFileNames.clear();

        // lấy địa chỉ file pdf
        pdfPath = filePDFPath;
        // lấy đi chỉ thư mục chứa file excel
//        csvExcelDirPath = fileCSVDirPath;
        // lấy đi chỉ thư mục chứa file excel csv
        csvExcelDirPath = fileChlDirPath;
        // lấy đi chỉ thư mục chứa chl
        chlDirPath = fileChlDirPath;

        // lấy mảng chứa các trang
        String[] kakuKouSyu = getFullToriaiText();
        // lấy trang đầu tiên và lấy ra các thông tin của đơn như tên khách hàng, ngày tháng
        getHeaderData(kakuKouSyu[0]);

        // chuyển mảng các trang sang dạng list
        List<String> kakuKouSyuList = new LinkedList<>(Arrays.asList(kakuKouSyu));

        // kích thước list
        int kakuKouSyuListSize = kakuKouSyuList.size();
        // lặp qua các trang gộp các trang cùng loại vật liệu làm 1 và xóa các trang đã được gộp vào trang khác đi
        for (int i = 1; i < kakuKouSyuListSize; i++) {
            // lấy tên vật liệu đang lặp
            String KouSyuName = extractValue(kakuKouSyuList.get(i), "法:", "梱包");

            // duyệt các trang phía sau, nếu vật liệu giống trang đang lặp thì gộp trang đó vào trang này
            // và xóa trang đó đi
            for (int j = i + 1; j < kakuKouSyuListSize; j++) {
                String KouSyuNameAfter = extractValue(kakuKouSyuList.get(j), "法:", "梱包");
                if (KouSyuName.equals(KouSyuNameAfter)) {
                    kakuKouSyuList.set(i, kakuKouSyuList.get(i).concat(kakuKouSyuList.get(j)));
                    kakuKouSyuList.remove(j);
                    j--;
                    kakuKouSyuListSize--;
                }
            }

            /*if (i > 1) {
                String KouSyuNameBefore = extractValue(kakuKouSyuList.get(i - 1), "法:", "梱包");

                if (KouSyuName.equals(KouSyuNameBefore)) {
                    kakuKouSyuList.set(i - 1, kakuKouSyuList.get(i - 1).concat(kakuKouSyuList.get(i)));
                    kakuKouSyuList.remove(i);
                    i--;
                    kakuKouSyuListSize--;
                }
            }*/
        }

        // đoạn code copy file này khác với app chl vì nó chỉ tạo 1 file nên chỉ chạy 1 lần ở đoạn đầu này
        // tạo path chứa file excel
        // mà không chạy trong vòng lặp bên dưới như trong hàm writeDataToChl
        excelPath = csvExcelDirPath + "\\" + fileExcelName + ".xlsx";
        // Tạo đối tượng File đại diện cho file cần xóa
        File file = new File(excelPath);
        // Kiểm tra nếu file tồn tại và xóa nó
        // vì nếu file đang được mở thì không thể ghi đè nhưng do file là readonly nên có thể xóa dù đang mở
        // xóa xong file thì có thể ghi lại file mới mà không bị lỗi không thể ghi đè
        if (file.exists()) {
            if (file.delete()) {
                System.out.println("File đã được xóa thành công.");
            } else {
                System.out.println("Xóa file thất bại.");
            }
        }
        // path chứa địa chỉ file sẽ được dán từ file copy
        Path copyFile = Paths.get(excelPath);
        // Đọc file mẫu từ resources rồi copy file ra địa chỉ của copyFile
        try (InputStream sourceFile = ReadPDFToExcel.class.getResourceAsStream("/com/example/convert_toriai_pdf_to_excel/sampleFiles/sample files.xlsx")) {
            if (sourceFile == null) {
                throw new IOException("File mẫu không tồn tại trong JAR ứng dụng");
            }
            Files.copy(sourceFile, copyFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
        // thêm tên file vào list các sheet của file để hiển thị tên file
        csvFileNames.add(new CsvFile("EXCEL: " + fileExcelName + ".xlsx", "", 0, 0));
/*        // Đặt quyền chỉ đọc cho file
        File readOnly = new File(excelPath);
        if (readOnly.exists()) {
            boolean result = readOnly.setReadOnly();
            if (result) {
                System.out.println("File is set to read-only.");
            } else {
                System.out.println("Failed to set file to read-only.");
            }
        } else {
            System.out.println("File does not exist.");
        }*/

        // lặp qua từng loại vật liệu trong list và ghi chúng vào các file excel
        for (int i = 1; i < kakuKouSyuList.size(); i++) {
            // tách các đoạn bozai thành mảng
            String[] kakuKakou = kakuKouSyuList.get(i).split("加工No:");

            // tại đoạn đầu tiên sẽ không chứa bozai mà chứa tên vật liệu
            // lấy ra thông số loại vật liệu và 3 size riêng lẻ của vật liệu
            getKouSyu(kakuKakou);
            // tạo map kaKouPairs và nhập thông tin tính vật liệu vào
            // kaKouPairs là map chứa key cũng là map chỉ có 1 cặp có key là chiều dài bozai, value là số lượng bozai
            // còn value của kaKouPairs cũng là map chứa các cặp key là mảng 2 phần tử gồm tên và chiều dài sản phẩm, value là số lượng sản phẩm
            Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs = getToriaiData(kakuKakou);

//            writeDataToExcel(kaKouPairs, i - 1, csvFileNames);
//            writeDataToCSV(kaKouPairs, i - 1, csvFileNames);
            // ghi thông tin vào file định dạng sysc2 là file của chl
//            writeDataToChl(kaKouPairs, i, csvFileNames);
            writeDataToExcelToriai(kaKouPairs, i, csvFileNames);
        }

    }

    /**
     * lấy toàn bộ text của file pdf
     *
     * @return mảng chứa các trang của file pdf, đầu trang chứa tên vật liệu
     */
    private static String[] getFullToriaiText() throws IOException {
        // khởi tạo mảng, có thể ko cần nếu sau đó nó có thể được gán bằng mảng khác
        String[] kakuKouSyu = new String[0];
        // dùng thư viện đọc file pdf lấy toàn bộ text của file
        try (PDDocument document = PDDocument.load(new File(pdfPath))) {
            // nếu file không được mã hóa thì mới lấy được text
            if (!document.isEncrypted()) {
                PDFTextStripper pdfStripper = new PDFTextStripper();
                String toriaiText = pdfStripper.getText(document);

                // chia thành các trang thông qua đoạn 材寸, mỗi trang sẽ chứa loại vật liệu ở đầu trang
                kakuKouSyu = toriaiText.split("材寸");

//                System.out.println(toriaiText);

            }
        }

        return kakuKouSyu;
    }

    /**
     * lấy các thông tin của đơn và ghi vào các biến nhớ toàn cục
     * các thông tin nằm trong vùng xác định, dùng hàm extractValue để lấy
     *
     * @param header text chứa thông tin
     */
    private static void getHeaderData(String header) {
        String nouKi = extractValue(header, "期[", "]");
        String[] nouKiArr = nouKi.split("/");
        shortNouKi = nouKiArr[0] + nouKiArr[1] + nouKiArr[2];

        bikou = extractValue(header, "考[", "]");
        kyakuSakiMei = extractValue(header, "客先名[", "]");
        String names = extractValue(header, "工事名[", "]");
        String[] namesArr = names.split("\\+");
        if (namesArr.length == 3) {
            fileExcelName = namesArr[0];
            chuyuBan = namesArr[1];
            teiHaiSha = namesArr[2];
        } else {
            fileExcelName = names;
        }

        System.out.println(shortNouKi + " : " + bikou + " : " + kyakuSakiMei + " : " + chuyuBan + " : " + teiHaiSha);
    }

    /**
     * lấy thông số đầy đủ của vật liệu, tên vật liệu, mã vật liệu, 3 size của vật liệu và ghi vào biến toàn cục
     *
     * @param kakuKakou mảng chứa các tính vật liệu của vật liệu đang xét
     */
    private static void getKouSyu(String[] kakuKakou) {

        // lấy loại vật liệu tại mảng 0 và tách mảng 0 thành các dòng rồi lấu dòng đầu tiên
        // tại dòng này lấy loại vật liệu trong đoạn "法:", "梱包"
        kouSyu = extractValue(kakuKakou[0].split("\n")[0], "法:", "梱包");
        // phân tách vật liệu thành các đoạn thông tin
        String[] kouSyuNameAndSize = kouSyu.split("-");
        // lấy tên vật liệu tại index 0
        kouSyuName = kouSyuNameAndSize[0].trim();

        // từ tên vật liệu lấy ra được  số đại diện cho nó
        switch (kouSyuName) {
            case "K":
                koSyuNumMark = "3";
                break;
            case "L":
                koSyuNumMark = "4";
                break;
            case "FB":
                koSyuNumMark = "5";
                break;
            case "[":
                koSyuNumMark = "6";
                break;
            case "C":
                koSyuNumMark = "7";
                break;
            case "H":
                koSyuNumMark = "8";
                break;
            case "CA":
                koSyuNumMark = "9";
                break;
        }

        // lấy đoạn thông tin 2 chứa các size của vật liệu và phân tách nó thành mảng chứa các size này
        String[] koSyuSizeArr = kouSyuNameAndSize[1].split("x");

        size1 = 0;
        size2 = 0;
        size3 = 0;

        // với từng loại vật liệu có số lượng size khác nhau thì sẽ ghi khác nhau, do chỉ cần thông tin của 3 size và x10
        // size thừa sẽ không cần ghi
        if (koSyuSizeArr.length == 3) {
            size1 = convertStringToIntAndMul(koSyuSizeArr[1], 10);
            size2 = convertStringToIntAndMul(koSyuSizeArr[0], 10);
            size3 = convertStringToIntAndMul(koSyuSizeArr[2], 10);
        } else if (koSyuSizeArr.length == 4) {
            size1 = convertStringToIntAndMul(koSyuSizeArr[1], 10);
            size2 = convertStringToIntAndMul(koSyuSizeArr[0], 10);
            size3 = convertStringToIntAndMul(koSyuSizeArr[3], 10);
        } else {
            size1 = convertStringToIntAndMul(koSyuSizeArr[1], 10);
            size2 = convertStringToIntAndMul(koSyuSizeArr[0], 10);
        }
    }

    /**
     * phân tích tính vật liệu của vật liệu đang xét và gán vào map thông tin
     *
     * @param kakuKakou mảng chứa các tính vật liệu của vật liệu đang xét
     * @return map các đoạn tính vật liệu chứa key cũng là map chỉ có 1 cặp có key là chiều dài bozai, value là số lượng bozai
     * còn value của kaKouPairs cũng là map chứa các cặp key là mảng 2 phần tử gồm tên và chiều dài sản phẩm, value là số lượng sản phẩm
     */
    private static Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> getToriaiData(String[] kakuKakou) throws TimeoutException {
        rowToriAiNum = 0;
        // tạo map
        Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs = new LinkedHashMap<>();

        // nếu không có thông tin thì thoát
        if (kakuKakou == null) {
            return kaKouPairs;
        }

        // lặp qua các đoạn bozai và thêm chúng vào map chứa toàn bộ thông tin vật liệu
        for (int i = 1; i < kakuKakou.length; i++) {
            // lấy kirirosu tại lần 1
            if (i == 1) {
                kirirosu = extractValue(kakuKakou[i], "切りﾛｽ設定:", "mm");
            }

            // lấy đoạn bozai đang lặp
            String kaKouText = kakuKakou[i];

            // map chứa cặp chiều dài, số lượng bozai
            Map<StringBuilder, Integer> kouZaiChouPairs = new LinkedHashMap<>();
            // map chứa cặp key là mảng chứa tên + chiều dài sản phẩm, value là số lượng
            Map<StringBuilder[], Integer> meiSyouPairs = new LinkedHashMap<>();

            // tạo mảng chứa các dòng trong đoạn bozai
            String[] kaKouLines = kaKouText.split("\n");

            // duyệt qua các dòng để thêm vào map
            for (String line : kaKouLines) {
                // nếu dòng có 鋼材長 và 本数 thì là dòng chứa bozai
                // lấy bozai và số lượng thêm vào map
                // mẫu định dạng "#.##". Mẫu này chỉ hiển thị phần thập phân nếu có, và tối đa là 2 chữ số thập phân.
                DecimalFormat df = new DecimalFormat("#.##");
                if (line.contains("鋼材長:") && line.contains("本数:")) {
                    String kouZaiChou = extractValue(line, "鋼材長:", "mm").trim();
                    String honSuu = extractValue(line, "本数:", " ").split(" ")[0].trim();

                    kouZaiChouPairs.put(new StringBuilder().append(df.format(Double.parseDouble(kouZaiChou))), convertStringToIntAndMul(honSuu, 1));
                }

                // nếu dòng chứa 名称 thì là dòng sản phẩm
                if (line.contains("名称")) {
                    // lấy vùng chứa tên và chiều dài sản phẩm
                    String meiSyouLength = extractValue(line, "名称", "mm x").trim();
                    // tách vùng trên thành mảng chứa các phần tử tên và chiều dài
                    String[] meiSyouLengths = meiSyouLength.split(" ");

                    // tạo biến chứa tên
                    String name = "";
                    // vì vùng chứa chiều dài có thể có dấu cách nên phải lấy từ phần tử đầu tiên đến phần tử trước phần tử cuối cùng
                    // và cuối tên sẽ không thêm dấu cách
                    for (int j = 0; j < meiSyouLengths.length - 1; j++) {
                        String namej = meiSyouLengths[j];
                        name = name.concat(namej + " ");
                    }
                    // xóa dấu cách ở 2 đầu
                    name = name.trim();

                    // lấy vùng chứa chiều dài là vùng cuối cùng trong mảng tên
                    String length = meiSyouLengths[meiSyouLengths.length - 1].trim();

                    double dLength = Double.parseDouble(length);
                    seiHinSet.add(dLength);

                    // thêm tên và chiều dài vào mảng, tên với ứng dụng này thì không cần
                    StringBuilder[] nameAndLength = {new StringBuilder(), new StringBuilder().append(df.format(dLength))};

                    // lấy số lượng sản phẩm
                    String meiSyouHonSuu = extractValue(line, "mm x", "本").trim();
                    // thêm cặp tên + chiều dài và số lượng vào map
                    meiSyouPairs.put(nameAndLength, convertStringToIntAndMul(meiSyouHonSuu, 1));
                }
            }

            // thêm 2 map chứa thông tin vật liệu vào map gốc
            kaKouPairs.put(kouZaiChouPairs, meiSyouPairs);
        }

        // xắp xếp lại seiHinSet
        // Convert LinkedHashSet to an ArrayList
        ArrayList<Double> array = new ArrayList<>(seiHinSet);
        // sort ArrayList
        Collections.sort(array);
        // xóa các phần tử của set và thêm lại bằng ArrayList đã xắp xếp
        seiHinSet.clear();
        seiHinSet.addAll(array);

//        System.out.println("số sản phẩm trong set: " + seiHinSet.size());
//        seiHinSet.forEach(aDouble -> {
//            System.out.println("num: " + aDouble);
//        });


        // in thông tin vật liệu
        kaKouPairs.forEach((kouZaiChouPairs, meiSyouPairs) -> {
            kouZaiChouPairs.forEach((key, value) -> System.out.println("\n" + key.toString() + " : " + value));
            meiSyouPairs.forEach((key, value) -> System.out.println(key[0].toString() + " " + key[1].toString() + " : " + value));
        });

        // lặp qua các phần tử của map kaKouPairs để tính số dòng sản phẩm đã lấy được
        for (Map.Entry<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> e : kaKouPairs.entrySet()) {

            // lấy map chiều dài bozai và số lượng
            Map<StringBuilder, Integer> kouZaiChouPairs = e.getKey();
            // lấy map tên + chiều dài sản phẩm và số lượng
            Map<StringBuilder[], Integer> meiSyouPairs = e.getValue();
            // tạo biến chứa số lượng bozai
            int kouZaiNum = 1;
            // lặp qua map bozai lấy giá trị số lượng bozai
            for (Map.Entry<StringBuilder, Integer> entry : kouZaiChouPairs.entrySet()) {
                kouZaiNum = entry.getValue();
            }

            // lấy kết quả số dòng sản phẩm đã lấy được bằng cách lấy số dòng của các lần lặp trước + số dòng của lần này(kouZaiNum * meiSyouPairs.size())
            // meiSyouPairs.size chính là số sản phẩm của bozai đang lặp
            rowToriAiNum += kouZaiNum * meiSyouPairs.size();
        }

        // nếu số dòng lớn hơn 99 th cho bằng 99 rồi ném ngoại lệ timeout để cho chương trình biết rồi hiển thị thông báo
        if (rowToriAiNum > 99) {
            rowToriAiNum = 99;
            System.out.println("vượt quá 99 hàng");
            // lấy tên file chl trong tiêu đề gắn thêm tên vật liệu + .sysc2 để in ra thông báo
            fileName = fileExcelName + " " + kouSyu + ".sysc2";
            throw new TimeoutException();
        }

        System.out.println(rowToriAiNum);
        System.out.println("\n" + kirirosu);

        // trả về map kết quả để ghi vào file chl sysc2
        return kaKouPairs;
    }

    private static void writeDataToExcel(Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs, int timePlus, ObservableList<CsvFile> csvFileNames) throws FileNotFoundException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        // Ghi thời gian hiện tại vào ô A1
        Row row1 = sheet.createRow(0);
        Cell cellA1 = row1.createCell(0);

        // Ghi thời gian hiện tại vào dòng đầu tiên
        Date currentDate = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat("yyMMddHHmm");
//        SimpleDateFormat sdfSecond = new SimpleDateFormat("yyMMddHHmmss");

        // Tăng thời gian lên timePlus phút
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(currentDate);
        calendar.add(Calendar.MINUTE, timePlus);

        // Lấy thời gian sau khi tăng
        Date newDate = calendar.getTime();

        String newTime = sdf.format(currentDate);

        cellA1.setCellValue(newTime + "+" + timePlus);

        // Ghi size1, size2, size3, 1 vào ô A2, B2, C2, D2
        Row row2 = sheet.createRow(1);
        row2.createCell(0).setCellValue(size1);
        row2.createCell(1).setCellValue(size2);
        row2.createCell(2).setCellValue(size3);
        row2.createCell(3).setCellValue(1);

        // Ghi koSyuNumMark, 1, rowToriAiNum, 1 vào ô A3, B3, C3, D3
        Row row3 = sheet.createRow(2);
        row3.createCell(0).setCellValue(koSyuNumMark);
        row3.createCell(1).setCellValue(1);
        row3.createCell(2).setCellValue(rowToriAiNum);
        row3.createCell(3).setCellValue(1);

        int rowIndex = 3;

        // tổng chiều dài các kozai
        double kouzaiChouGoukei = 0;
        double seiHinChouGoukei = 0;
        // Ghi dữ liệu từ KA_KOU_PAIRS vào các ô
        for (Map.Entry<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> entry : kaKouPairs.entrySet()) {
            if (rowIndex >= 102) break;

            Map<StringBuilder, Integer> kouZaiChouPairs = entry.getKey();
            Map<StringBuilder[], Integer> meiSyouPairs = entry.getValue();

            String keyTemp = "";
            int valueTemp = 0;

            // Ghi dữ liệu từ mapkey vào ô D4
            for (Map.Entry<StringBuilder, Integer> kouZaiEntry : kouZaiChouPairs.entrySet()) {

                keyTemp = String.valueOf(kouZaiEntry.getKey());
                valueTemp = kouZaiEntry.getValue();
                // cộng thêm chiều dài của bozai * số lượng vào tổng
                kouzaiChouGoukei += Double.parseDouble(keyTemp) * valueTemp;
            }

            // Ghi dữ liệu từ mapvalue vào ô A4, B4 và các hàng tiếp theo
            for (int i = 0; i < valueTemp; i++) {
                int j = 0;
                for (Map.Entry<StringBuilder[], Integer> meiSyouEntry : meiSyouPairs.entrySet()) {
                    if (rowIndex >= 102) break;
                    // chiều dài sản phẩm
                    String leng = String.valueOf(meiSyouEntry.getKey()[1]);
                    // số lượng sản phẩm
                    String num = meiSyouEntry.getValue().toString();

                    Row row = sheet.createRow(rowIndex++);
                    row.createCell(0).setCellValue(leng);
                    row.createCell(1).setCellValue(num);
                    row.createCell(2).setCellValue(String.valueOf(meiSyouEntry.getKey()[0]));

                    // cộng thêm vào chiều dài của sản phẩm * số lượng vào tổng
                    seiHinChouGoukei += Double.parseDouble(leng) * Double.parseDouble(num);
                    j++;
                }
                sheet.getRow(rowIndex - j).createCell(3).setCellValue(keyTemp);
            }
        }

/*        // không cần tạo nữa vì chiều dài bozai sẽ ghi vào cột 4
        // thay vì cột 3 như trước nên không thể ghi thêm các thông tin này vào cột 4 nữa
        // nếu không có hàng sản phẩm nào thì sẽ chưa tạo hàng 4, 5, 6, 7, 8 và rowIndex vẫn là 3
        // cần tạo thêm 4 hàng này để ghi các thông tin kouJiMe, kyakuSakiMei, shortNouKi, kirirosu, fileName bên dưới
        for (int i = 0; i < 5; i++) {
            if (rowIndex <= i + 3) {
                sheet.createRow(i + 3);
            }
        }

        // Ghi kouJiMe, kyakuSakiMei, shortNouKi, kirirosu, fileName + kouSyu vào ô D4, D5, D6, D7
        sheet.getRow(3).createCell(3).setCellValue(kouJiMe);
        sheet.getRow(4).createCell(3).setCellValue(kyakuSakiMei);
        sheet.getRow(5).createCell(3).setCellValue(shortNouKi);
        sheet.getRow(6).createCell(3).setCellValue(kirirosu);
        sheet.getRow(7).createCell(3).setCellValue(fileChlName + " " + kouSyu);*/

        // Ghi giá trị 0 vào các ô A99, B99, C99, D99
        Row lastRow = sheet.createRow(rowIndex);
        lastRow.createCell(0).setCellValue(0);
        lastRow.createCell(1).setCellValue(0);
        lastRow.createCell(2).setCellValue(0);
        lastRow.createCell(3).setCellValue(0);

        String[] linkarr = pdfPath.split("\\\\");
//        fileName = linkarr[linkarr.length - 1].split("\\.")[0] + " " + kouSyu + ".xlsx";
        fileName = fileExcelName + " " + kouSyu + ".xlsx";
//        String fileNameAndTime = linkarr[linkarr.length - 1].split("\\.")[0] + "(" + sdfSecond.format(currentDate) + ")--" + kouSyu + ".csv";
        String excelPath = csvExcelDirPath + "\\" + fileName;

        // Tạo đối tượng File đại diện cho file cần xóa
        File file = new File(excelPath);

        // Kiểm tra nếu file tồn tại và xóa nó
        // vì nếu file đang được mở thì không thể ghi đè nhưng do file là readonly nên có thể xóa dù đang mở
        // xóa xong file thì có thể ghi lại file mới mà không bị lỗi không thể ghi đè
        if (file.exists()) {
            if (file.delete()) {
//                System.out.println("File đã được xóa thành công.");
            } else {
//                System.out.println("Xóa file thất bại.");
            }
        }

        try (FileOutputStream fileOut = new FileOutputStream(excelPath)) {
            workbook.write(fileOut);
            workbook.close();
        } catch (IOException e) {
            if (e instanceof FileNotFoundException) {
                System.out.println("File đang được mở bởi người dùng khác");
                throw new FileNotFoundException();
            }
            System.out.println(e.getMessage());
            throw new RuntimeException(e);
        }

        // Đặt quyền chỉ đọc cho file
        File readOnly = new File(excelPath);
        if (readOnly.exists()) {
            boolean result = readOnly.setReadOnly();
            if (result) {
                System.out.println("File is set to read-only.");
            } else {
                System.out.println("Failed to set file to read-only.");
            }
        } else {
            System.out.println("File does not exist.");
        }

        System.out.println("tong chieu dai bozai " + kouzaiChouGoukei);
        System.out.println("tong chieu dai san pham " + seiHinChouGoukei);
        csvFileNames.add(new CsvFile(fileName, kouSyuName, kouzaiChouGoukei, seiHinChouGoukei));

    }

    private static void writeDataToCSV(Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs, int timePlus, ObservableList<CsvFile> csvFileNames) throws FileNotFoundException {

        // Ghi thời gian hiện tại vào dòng đầu tiên
        Date currentDate = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat("yyMMddHHmm");
//        // Tạo thêm fomat có thêm giây
//        SimpleDateFormat sdfSecond = new SimpleDateFormat("yyMMddHHmmss");

        /*// Tăng thời gian lên timePlus phút
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(currentDate);
        calendar.add(Calendar.MINUTE, timePlus);

        // Lấy thời gian sau khi tăng
        Date newDate = calendar.getTime();

        String newTime = sdf.format(newDate);*/

        // lấy thời gian hiện tại với fomat đã chọn
        String currentTime = sdf.format(currentDate);

        String[] linkarr = pdfPath.split("\\\\");
//        fileName = linkarr[linkarr.length - 1].split("\\.")[0] + " " + kouSyu + ".csv";
        fileName = fileExcelName + " " + kouSyu + ".csv";
//        // tạo tên file có gắn thêm thời gian để không trùng với file trước đó
//        String fileNameAndTime = linkarr[linkarr.length - 1].split("\\.")[0] + "(" + sdfSecond.format(currentDate) + ")--" + kouSyu + ".csv";
        String csvPath = csvExcelDirPath + "\\" + fileName;
        System.out.println("dir path: " + csvExcelDirPath);
        System.out.println("filename: " + fileName);

        // Tạo đối tượng File đại diện cho file cần xóa
        File file = new File(csvPath);

        // Kiểm tra nếu file tồn tại và xóa nó
        // vì nếu file đang được mở thì không thể ghi đè nhưng do file là readonly nên có thể xóa dù đang mở
        // xóa xong file thì có thể ghi lại file mới mà không bị lỗi không thể ghi đè
        if (file.exists()) {
            if (file.delete()) {
//                System.out.println("File đã được xóa thành công.");
            } else {
//                System.out.println("Xóa file thất bại.");
            }
        } else {
//            System.out.println("File không tồn tại.");
        }
        // tổng chiều dài các kozai
        double kouzaiChouGoukei = 0;
        double seiHinChouGoukei = 0;
        try (CSVWriter writer = new CSVWriter(new OutputStreamWriter(new FileOutputStream(csvPath), Charset.forName("MS932")))) {


            writer.writeNext(new String[]{currentTime + "+" + timePlus});

            // Ghi size1, size2, size3, 1 vào dòng tiếp theo
            writer.writeNext(new String[]{String.valueOf(size1), String.valueOf(size2), String.valueOf(size3), "1"});

            // Ghi koSyuNumMark, 1, rowToriAiNum, 1 vào dòng tiếp theo
            writer.writeNext(new String[]{koSyuNumMark, "1", String.valueOf(rowToriAiNum), "1"});

            List<String[]> toriaiDatas = new LinkedList<>();

            int rowIndex = 3;

            // Ghi dữ liệu từ KA_KOU_PAIRS vào các ô
            for (Map.Entry<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> entry : kaKouPairs.entrySet()) {
                if (rowIndex >= 102) break;

                Map<StringBuilder, Integer> kouZaiChouPairs = entry.getKey();
                Map<StringBuilder[], Integer> meiSyouPairs = entry.getValue();

                String keyTemp = "";
                int valueTemp = 0;

                // Ghi dữ liệu từ mapkey vào ô D4
                for (Map.Entry<StringBuilder, Integer> kouZaiEntry : kouZaiChouPairs.entrySet()) {

                    keyTemp = String.valueOf(kouZaiEntry.getKey());
                    valueTemp = kouZaiEntry.getValue();

                    // cộng thêm chiều dài của bozai * số lượng vào tổng
                    kouzaiChouGoukei += Double.parseDouble(keyTemp) * valueTemp;
                }

                // Ghi dữ liệu từ mapvalue vào ô A4, B4 và các hàng tiếp theo
                for (int i = 0; i < valueTemp; i++) {
                    int j = 0;
                    for (Map.Entry<StringBuilder[], Integer> meiSyouEntry : meiSyouPairs.entrySet()) {
                        if (rowIndex >= 102) break;

                        String[] line = new String[4];
                        rowIndex++;

                        // chiều dài sản phẩm
                        String leng = String.valueOf(meiSyouEntry.getKey()[1]);
                        // số lượng sản phẩm
                        String num = meiSyouEntry.getValue().toString();
                        // ghi chiều dài sản phẩm
                        line[0] = leng;
                        // ghi số lượng sản phẩm
                        line[1] = num;
                        line[2] = String.valueOf(meiSyouEntry.getKey()[0]);

                        // cộng thêm vào chiều dài của sản phẩm * số lượng vào tổng
                        seiHinChouGoukei += Double.parseDouble(leng) * Double.parseDouble(num);
                        toriaiDatas.add(line);
                        j++;
                    }
                    toriaiDatas.get(toriaiDatas.size() - j)[3] = keyTemp;
                }
            }

/*            // không cần tạo nữa vì chiều dài bozai sẽ ghi vào cột 4
            // thay vì cột 3 như trước nên không thể ghi thêm các thông tin này vào cột 4 nữa
            // nếu không có hàng sản phẩm nào thì sẽ chưa tạo hàng 4, 5, 6, 7, 8 và rowIndex vẫn là 3
            // cần tạo thêm 4 hàng này để ghi các thông tin kouJiMe, kyakuSakiMei, shortNouKi, kirirosu, fileName bên dưới
            for (int i = 0; i < 5; i++) {
                if (rowIndex <= i + 3) {
                    toriaiDatas.add(new String[4]);
                }
            }

            // Ghi kouJiMe, kyakuSakiMei, shortNouKi, kirirosu, fileName + " " + kouSyu vào ô D4, D5, D6, D7
            toriaiDatas.get(0)[3] = kouJiMe;
            toriaiDatas.get(1)[3] = kyakuSakiMei;
            toriaiDatas.get(2)[3] = shortNouKi;
            toriaiDatas.get(3)[3] = kirirosu;
            toriaiDatas.get(4)[3] = fileChlName + " " + kouSyu;*/

            writer.writeAll(toriaiDatas);

            // Ghi giá trị 0 vào các ô A99, B99, C99, D99
            writer.writeNext(new String[]{"0", "0", "0", "0"});


        } catch (IOException e) {
            if (e instanceof FileNotFoundException) {
                System.out.println("File đang được mở bởi người dùng khác");
                throw new FileNotFoundException();
            }
            System.out.println(e.getMessage());
            throw new RuntimeException(e);
        }

        // Đặt quyền chỉ đọc cho file
        File readOnly = new File(csvPath);
        if (readOnly.exists()) {
            boolean result = readOnly.setReadOnly();
            if (result) {
//                System.out.println("File is set to read-only.");
            } else {
//                System.out.println("Failed to set file to read-only.");
            }
        } else {
//            System.out.println("File does not exist.");
        }

        System.out.println("tong chieu dai bozai " + kouzaiChouGoukei);
        System.out.println("tong chieu dai san pham " + seiHinChouGoukei);
        csvFileNames.add(new CsvFile(fileName, kouSyuName, kouzaiChouGoukei, seiHinChouGoukei));

    }

    /**
     * ghi tính vật liệu của vật liệu đang xét trong map vào file mới
     *
     * @param kaKouPairs   map chứa tính vật liệu
     * @param timePlus     thời gian hoặc chỉ số cộng thêm vào ô time để tránh bị trùng tên  time giữa các file
     * @param csvFileNames list chứa danh sách các file đã tạo
     */
    private static void writeDataToChl(Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs, int timePlus, ObservableList<CsvFile> csvFileNames) throws FileNotFoundException {

        // Ghi thời gian hiện tại vào dòng đầu tiên
        Date currentDate = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat("yyMMddHHmm");
//        // Tạo thêm fomat có thêm giây
//        SimpleDateFormat sdfSecond = new SimpleDateFormat("yyMMddHHmmss");

/*        // Tăng thời gian lên timePlus phút
        // hiện tại không dùng đoạn code này nữa
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(currentDate);
        calendar.add(Calendar.MINUTE, timePlus);

        // Lấy thời gian sau khi tăng
        Date newDate = calendar.getTime();

        String newTime = sdf.format(newDate);*/

        // lấy thời gian hiện tại với fomat đã chọn
        String currentTime = sdf.format(currentDate);

        // lấy tên file chl trong tiêu đề gắn thêm tên vật liệu + .sysc2
        fileName = fileExcelName + " " + kouSyu + ".sysc2";

//        // tạo tên file có gắn thêm thời gian để không trùng với file trước đó
//        String fileNameAndTime = linkarr[linkarr.length - 1].split("\\.")[0] + "(" + sdfSecond.format(currentDate) + ")--" + kouSyu + ".csv";

        String chlPath = chlDirPath + "\\" + fileName;
        System.out.println("dir path: " + csvExcelDirPath);
        System.out.println("filename: " + fileName);


        // Tạo đối tượng File đại diện cho file cần xóa
        File file = new File(chlPath);

        // Kiểm tra nếu file tồn tại và xóa nó
        // vì nếu file đang được mở thì không thể ghi đè nhưng do file là readonly nên có thể xóa dù đang mở
        // xóa xong file thì có thể ghi lại file mới mà không bị lỗi không thể ghi đè
        if (file.exists()) {
            if (file.delete()) {
//                System.out.println("File đã được xóa thành công.");
            } else {
//                System.out.println("Xóa file thất bại.");
            }
        } else {
//            System.out.println("File không tồn tại.");
        }

        // tổng chiều dài các kozai
        double kouzaiChouGoukei = 0;
        double seiHinChouGoukei = 0;
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(chlPath, Charset.forName("MS932")))) {

            writer.write(currentTime + "+" + timePlus + ",,,");
            writer.newLine();


            // Ghi size1, size2, size3, 1 vào dòng tiếp theo
            writer.write(size1 + "," + size2 + "," + size3 + "," + "1");
            writer.newLine();

            // Ghi koSyuNumMark, 1, 99, 1 vào dòng tiếp theo, rowToriAiNum sẽ được sử dụng sau khi ước tính ghi đến hàng 102
            writer.write(koSyuNumMark + "," + "1" + "," + "99" + "," + "1");
            writer.newLine();

            // tạo list chứa các mảng, mỗi mảng là 1 dòng cần ghi theo fomat của chl
            List<String[]> toriaiDatas = new LinkedList<>();

            int rowIndex = 3;

            // Ghi dữ liệu từ KA_KOU_PAIRS vào các ô
            // kaKouPairs là map chứa key cũng là map chỉ có 1 cặp có key là chiều dài bozai, value là số lượng bozai
            // còn value của kaKouPairs cũng là map chứa các cặp key là tên + chiều dài sản phẩm, value là số lượng sản phẩm
            for (Map.Entry<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> entry : kaKouPairs.entrySet()) {
                if (rowIndex >= 102) break;

                Map<StringBuilder, Integer> kouZaiChouPairs = entry.getKey();
                Map<StringBuilder[], Integer> meiSyouPairs = entry.getValue();

                // chiều dài bozai
                String keyTemp = "";
                // số lượng bozai
                int valueTemp = 0;


                // Ghi dữ liệu bozai từ mapkey vào ô D4 kouZaiChouPairs
                for (Map.Entry<StringBuilder, Integer> kouZaiEntry : kouZaiChouPairs.entrySet()) {
                    keyTemp = String.valueOf(kouZaiEntry.getKey());
                    valueTemp = kouZaiEntry.getValue();
                    // cộng thêm chiều dài của bozai * số lượng vào tổng
                    kouzaiChouGoukei += Double.parseDouble(keyTemp) * valueTemp;
                }

                // Ghi dữ liệu từ mapvalue vào ô A4, B4 và các hàng tiếp theo
                // số lượng bozai là bao nhiêu thì phải ghi bấy nhiêu lần
                for (int i = 0; i < valueTemp; i++) {
                    int j = 0; // đếm số hàng đã ghi
                    // lặp qua map sản phẩm, tính chiều dài map bằng j
                    for (Map.Entry<StringBuilder[], Integer> meiSyouEntry : meiSyouPairs.entrySet()) {
                        if (rowIndex >= 102) break;

                        // tạo mảng lưu dòng đang lặp gồm 4 phần tử lần lượt là
                        // chiều dài sản phẩm, số lượng sản phẩm, tên sản phẩm, chiều dài bozai
                        String[] line = new String[4];
                        rowIndex++;

                        // chiều dài sản phẩm
                        String leng = String.valueOf(meiSyouEntry.getKey()[1]);
                        // số lượng sản phẩm
                        String num = meiSyouEntry.getValue().toString();
                        // ghi chiều dài sản phẩm
                        line[0] = leng;
                        // ghi số lượng sản phẩm
                        line[1] = num;
                        // ghi tên sản phẩm
                        line[2] = String.valueOf(meiSyouEntry.getKey()[0]);
                        // ghi vào phần tử thứ 3 của mảng giá trị rỗng để tránh giá trị null
                        line[3] = "";

                        // cộng thêm vào chiều dài của sản phẩm * số lượng vào tổng
                        seiHinChouGoukei += Double.parseDouble(leng) * Double.parseDouble(num);

                        // thêm hàng sản phẩm vừa tạo vào list
                        toriaiDatas.add(line);
                        // tăng số hàng lên 1
                        j++;
                    }
                    // ghi vào cột 4 ([3]) chiều dài bozai khi ghi xong 1 lượt sản phẩm + số lượng
                    // tính vị trí của nó bằng cách lấy size của list kaKouPairs - chiều dài map sản phẩm
                    toriaiDatas.get(toriaiDatas.size() - j)[3] = keyTemp;
                }
            }

/*
            // nếu không có hàng sản phẩm nào thì sẽ chưa tạo hàng 4, 5, 6, 7, 8 và rowIndex vẫn là 3
            // cần tạo thêm 4 hàng này để ghi các thông tin kouJiMe, kyakuSakiMei, shortNouKi, kirirosu, fileName bên dưới
            // không cần tạo nữa vì ghi file sysc2 sẽ ghi xuống cuối
            for (int i = 0; i < 5; i++) {
                if (rowIndex <= i + 3) {
                    toriaiDatas.add(new String[4]);
                }
            }
*/

/*
            // Ghi kouJiMe, kyakuSakiMei, shortNouKi, kirirosu, fileName + " " + kouSyu vào ô D4, D5, D6, D7
            // không cần tạo nữa vì ghi file sysc2 sẽ ghi xuống cuối và vì chiều dài bozai sẽ ghi vào cột 4
            // thay vì cột 3 như trước nên không thể ghi thêm các thông tin này vào cột 4 nữa
            toriaiDatas.get(0)[3] = kouJiMe;
            toriaiDatas.get(1)[3] = kyakuSakiMei;
            toriaiDatas.get(2)[3] = shortNouKi;
            toriaiDatas.get(3)[3] = kirirosu;
            toriaiDatas.get(4)[3] = fileChlName + " " + kouSyu;
*/
            // lặp qua list chứa các dòng toriaiDatas
            for (String[] line : toriaiDatas) {

/*                // cách ghi này không dùng được nữa vì cách ghi phần tử cuối cùng đã thay đổi
                for (String length : line) {
                    writer.write(length + ",");
                }*/

                // mỗi dòng là 1 mảng nên lặp qua mảng ghi các phần tử vào dòng phân tách nhau bởi dấu (,)
                for (int i = 0; i < line.length; i++) {
                    if (i == line.length - 1) {
                        writer.write(line[i]);
                    } else {
                        writer.write(line[i] + ",");
                    }
                }
                writer.newLine();
            }

            // ghi nốt các dòng còn lại không có sản phẩn ",,," để đủ 99 sản phẩm
            for (int i = toriaiDatas.size(); i < 99; i++) {
                writer.write(",,,");
                writer.newLine();
            }


            // Ghi giá trị 0 vào dòng tiếp theo là dòng 103
            writer.write("0,0,0,0");
            writer.newLine();
            // ghi 20 và kirirosu vào dòng tiếp
            writer.write("20.0," + kirirosu + ",,");
            writer.newLine();
            // ghi các tên và ngày vào dòng tiếp
            writer.write(bikou + "," + kyakuSakiMei + "," + shortNouKi + ",");
            writer.newLine();
            // dòng tiếp theo là ghi 備考１、備考２ theo định dạng 備考１,備考２,, nhưng không có nên không cần chỉ ghi (,,,)
            writer.write(",,,");
            writer.newLine();
            // ghi dấu hiệu nhận biết kết thúc
            writer.write("END,,,");
            writer.newLine();


        } catch (IOException e) {
            if (e instanceof FileNotFoundException) {
                System.out.println("File đang được mở bởi người dùng khác");
                throw new FileNotFoundException();
            }
            System.out.println(e.getMessage());
            throw new RuntimeException(e);
        }

        // Đặt quyền chỉ đọc cho file
        File readOnly = new File(chlPath);
        if (readOnly.exists()) {
            boolean result = readOnly.setReadOnly();
            if (result) {
//                System.out.println("File is set to read-only.");
            } else {
//                System.out.println("Failed to set file to read-only.");
            }
        } else {
//            System.out.println("File does not exist.");
        }

        System.out.println("tong chieu dai bozai " + kouzaiChouGoukei);
        System.out.println("tong chieu dai san pham " + seiHinChouGoukei);
        // thêm file vào list hiển thị
        csvFileNames.add(new CsvFile(fileName, kouSyuName, kouzaiChouGoukei, seiHinChouGoukei));

    }

    /**
     * trả về đoạn text nằm giữa startDelimiter và endDelimiter
     *
     * @param text           đoạn văn bản chứa thông tin tìm kiếm
     * @param startDelimiter đoạn text phía trước vùng cần tìm
     * @param endDelimiter   đoạn text phía sau vùng cần tìm
     * @return đoạn text nằm giữa startDelimiter và endDelimiter
     */
    private static String extractValue(String text, String startDelimiter, String endDelimiter) {
        // lấy index của startDelimiter + độ dài của nó để bỏ qua nó và xác định được index bắt đầu của đoạn text nó bao ngoài, chính là đoạn text cần tìm
        int startIndex = text.indexOf(startDelimiter) + startDelimiter.length();
        // lấy index của endDelimiter bắt đầu tìm từ index của startDelimiter để tránh tìm kiếm trong các vùng khác phía trước không liên quan, đây chính là
        // index cuối cùng của đoạn text cần tìm
        int endIndex = text.indexOf(endDelimiter, startIndex);
//        System.out.println(text);
        // trả về đoạn text cần tìm bằng 2 index vừa xác định ở trên
        return text.substring(startIndex, endIndex).trim();
    }


    private static void writeDataToExcelToriai(Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs, int sheetIndex, ObservableList<CsvFile> csvFileNames) throws FileNotFoundException {

        // tạo luồng đọc ghi file
        try (FileInputStream file = new FileInputStream(excelPath)) {
            Workbook workbook = new XSSFWorkbook(file);

            // nếu tên vật liệu có chứa [ thì phải đổi sang U vì tên này sẽ đặt tên cho sheet nên [ không dùng được
            if (kouSyu.contains("[")) {
                kouSyu = kouSyu.replace("[", "U");
            }

            // Lấy index sheet gốc cần sao chép
            int sheetSampleIndex = 0;
            // sao chép sheet gốc sang một sheet mới
            workbook.cloneSheet(sheetSampleIndex);
            // đổi tên sheet mới theo tên vật liệu đang duyệt, sheetIndex là chỉ số của sheet mới
            workbook.setSheetName(sheetIndex, kouSyu);
            // lấy ra sheet mới
            Sheet sheet = workbook.getSheetAt(sheetIndex);


            Date currentDate = new Date();
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");

            String time = sdf.format(currentDate);
            // Ghi thời gian hiện tại vào ô C1
            sheet.getRow(0).getCell(2).setCellValue(time);

            // Ghi tên khách hàng vào ô G6
            sheet.getRow(0).getCell(6).setCellValue(kyakuSakiMei);

            // Ghi bikou vào ô M12
            sheet.getRow(0).getCell(12).setCellValue(bikou);

            // Ghi shortNouKi vào ô S18
            sheet.getRow(0).getCell(18).setCellValue(shortNouKi);

            // Ghi saizu vào ô C2, chưa dùng
            sheet.getRow(1).getCell(2).setCellValue("");

            // Ghi chuyuBan vào ô I8
            sheet.getRow(1).getCell(8).setCellValue(chuyuBan);
            // Ghi teiHaiSha vào ô O14
            sheet.getRow(1).getCell(14).setCellValue(teiHaiSha);

            int soBoZai = kaKouPairs.size();
            int soSanPham = seiHinSet.size();

            if (soBoZai > 15) {

            }


            sheet.shiftColumns(4, sheet.getLastRowNum(), 1);
            sheet.shiftColumns(4 + 23, sheet.getLastRowNum(), 1);
            sheet.shiftColumns(4 + 41, sheet.getLastRowNum(), 1);

            for (int i = 0; i < 3; i++) {
                Row row = sheet.getRow(i);
                row.shiftCellsLeft(5, row.getLastCellNum(), 1);
            }

            for (int i = 26 + 0; i <= 41 + 0; i++) {
                Row row = sheet.getRow(6);
                Cell cell = row.getCell(i);

                if (cell != null && cell.getCellType() == CellType.FORMULA) {
                    String formula = cell.getCellFormula();
                    formula = formula.replaceAll("L", "K");
                    cell.setCellFormula(formula);
                }
            }

            Cell srcCell;
            Cell destCell;
            for (int i = 3; i <= 9; i++) {
                Row row = sheet.getRow(i);
                // Sao chép ô từ cột srcColumn sang destColumn
                srcCell = row.getCell(3);
                destCell = row.createCell(4);
                copyCellWithFormulaUpdate(srcCell, destCell, 1);
            }

            Row row7Formula = sheet.getRow(6);
            srcCell = row7Formula.getCell(26);
            destCell = row7Formula.createCell(27);
            copyCellWithFormulaUpdate(srcCell, destCell, 1);

            srcCell = row7Formula.getCell(44);
            destCell = row7Formula.createCell(45);
            copyCellWithFormulaUpdate(srcCell, destCell, 1);



            /*
            // Ghi koSyuNumMark, 1, rowToriAiNum, 1 vào ô A3, B3, C3, D3
            Row row3 = sheet.createRow(2);
            row3.createCell(0).setCellValue(koSyuNumMark);
            row3.createCell(1).setCellValue(1);
            row3.createCell(2).setCellValue(rowToriAiNum);
            row3.createCell(3).setCellValue(1);

            int rowIndex = 3;

            // tổng chiều dài các kozai
            double kouzaiChouGoukei = 0;
            double seiHinChouGoukei = 0;
            // Ghi dữ liệu từ KA_KOU_PAIRS vào các ô
            for (Map.Entry<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> entry : kaKouPairs.entrySet()) {
                if (rowIndex >= 102) break;

                Map<StringBuilder, Integer> kouZaiChouPairs = entry.getKey();
                Map<StringBuilder[], Integer> meiSyouPairs = entry.getValue();

                String keyTemp = "";
                int valueTemp = 0;

                // Ghi dữ liệu từ mapkey vào ô D4
                for (Map.Entry<StringBuilder, Integer> kouZaiEntry : kouZaiChouPairs.entrySet()) {

                    keyTemp = String.valueOf(kouZaiEntry.getKey());
                    valueTemp = kouZaiEntry.getValue();
                    // cộng thêm chiều dài của bozai * số lượng vào tổng
                    kouzaiChouGoukei += Double.parseDouble(keyTemp) * valueTemp;
                }

                // Ghi dữ liệu từ mapvalue vào ô A4, B4 và các hàng tiếp theo
                for (int i = 0; i < valueTemp; i++) {
                    int j = 0;
                    for (Map.Entry<StringBuilder[], Integer> meiSyouEntry : meiSyouPairs.entrySet()) {
                        if (rowIndex >= 102) break;
                        // chiều dài sản phẩm
                        String leng = String.valueOf(meiSyouEntry.getKey()[1]);
                        // số lượng sản phẩm
                        String num = meiSyouEntry.getValue().toString();

                        Row row = sheet.createRow(rowIndex++);
                        row.createCell(0).setCellValue(leng);
                        row.createCell(1).setCellValue(num);
                        row.createCell(2).setCellValue(String.valueOf(meiSyouEntry.getKey()[0]));

                        // cộng thêm vào chiều dài của sản phẩm * số lượng vào tổng
                        seiHinChouGoukei += Double.parseDouble(leng) * Double.parseDouble(num);
                        j++;
                    }
                    sheet.getRow(rowIndex - j).createCell(3).setCellValue(keyTemp);
                }
            }*/




            // Khóa sheet với mật khẩu
            sheet.protectSheet("");
            try (FileOutputStream fileOut = new FileOutputStream(excelPath)) {
                workbook.write(fileOut);

                workbook.close();
            }


        } catch (IOException e) {
            if (e instanceof FileNotFoundException) {
                System.out.println("File đang được mở bởi người dùng khác");
                throw new FileNotFoundException();
            }
            System.out.println(e.getMessage());
            throw new RuntimeException(e);
        }


//        System.out.println("tong chieu dai bozai " + kouzaiChouGoukei);
//        System.out.println("tong chieu dai san pham " + seiHinChouGoukei);
        csvFileNames.add(new CsvFile("Sheet " + sheetIndex + ": " + kouSyu, kouSyuName, 0, 0));

    }

    private static void copyCellWithFormulaUpdate(Cell srcCell, Cell destCell, int shiftColumns) {
        destCell.setCellStyle(srcCell.getCellStyle());
        switch (srcCell.getCellType()) {
            case STRING:
                destCell.setCellValue(srcCell.getStringCellValue());
                break;
            case NUMERIC:
                destCell.setCellValue(srcCell.getNumericCellValue());
                break;
            case BOOLEAN:
                destCell.setCellValue(srcCell.getBooleanCellValue());
                break;
            case FORMULA:
                String formula = srcCell.getCellFormula();
                String updatedFormula = updateFormula(formula, shiftColumns, srcCell.getRowIndex());
                updatedFormula = updatedFormula.replaceAll("SUN", "SUM");
                destCell.setCellFormula(updatedFormula);
                break;
            case BLANK:
                destCell.setBlank();
                break;
            default:
                break;
        }
    }

    private static String updateFormula(String formula, int shiftColumns, int rowIndex) {
        StringBuilder updatedFormula = new StringBuilder();
        int length = formula.length();

        for (int i = 0; i < length; i++) {
            char c = formula.charAt(i);
            if (Character.isLetter(c) || c == '$') {
                StringBuilder reference = new StringBuilder();
                boolean isColumnAbsolute = false;
                boolean isRowAbsolute = false;

                if (c == '$') {
                    isColumnAbsolute = true;
                    reference.append(c);
                    i++;
                    c = formula.charAt(i);
                }

                while (i < length && Character.isLetter(formula.charAt(i))) {
                    reference.append(formula.charAt(i));
                    i++;
                }

                if (i < length && formula.charAt(i) == '$') {
                    isRowAbsolute = true;
                    reference.append(formula.charAt(i));
                    i++;
                }

                while (i < length && Character.isDigit(formula.charAt(i))) {
                    reference.append(formula.charAt(i));
                    i++;
                }

                String column = reference.toString().replaceAll("[^A-Z]", "");
                String row = reference.toString().replaceAll("[^0-9]", "");

                if (!isColumnAbsolute) {
                    int columnIndex = columnToIndex(column) + shiftColumns;
                    updatedFormula.append(indexToColumn(columnIndex));
                } else {
                    updatedFormula.append(column);
                }

                if (!isRowAbsolute && !row.isEmpty()) {
                    updatedFormula.append(row);
                } else {
                    updatedFormula.append(row);
                }

                i--; // Adjust for the increment in the loop
            } else {
                updatedFormula.append(c);
            }
        }
        return updatedFormula.toString();
    }

    private static int columnToIndex(String column) {
        int index = 0;
        for (int i = 0; i < column.length(); i++) {
            index = index * 26 + (column.charAt(i) - 'A' + 1);
        }
        return index - 1;
    }

    private static String indexToColumn(int index) {
        StringBuilder column = new StringBuilder();
        while (index >= 0) {
            column.insert(0, (char) ('A' + (index % 26)));
            index = index / 26 - 1;
        }
        return column.toString();
    }


    /**
     * chuyển đổi text nhập vào sang số double rồi nhân với hệ số và trả về với kiểu int
     *
     * @param textNum    text cần chuyển
     * @param multiplier hệ số
     * @return số int đã nhân với hệ số
     */
    private static int convertStringToIntAndMul(String textNum, int multiplier) {
        Double num = null;
        try {
            num = Double.parseDouble(textNum);
        } catch (NumberFormatException e) {
            System.out.println("Lỗi chuyển đổi chuỗi không phải số thực sang số");
            System.out.println(textNum);

        }
        if (num != null) {
            return (int) (num * multiplier);
        }
        return 0;
    }
}

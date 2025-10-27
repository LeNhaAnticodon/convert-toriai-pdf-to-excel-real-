package com.example.convert_toriai_pdf_to_excel.dao;

import com.example.convert_toriai_pdf_to_excel.model.ExcelFile;
import com.example.convert_toriai_pdf_to_excel.model.Setup;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

public class SetupData {
    // biến thể hiện duy nhất của class
    private static SetupData instance;
    // tên file lưu cài đặt của chương trình
    private static String FILE_SETUP_NAME = "setup_data.set";
    // list các file đã chuyển sang chl
    private final ObservableList<ExcelFile> excelFiles = FXCollections.observableArrayList();
    // map chứa key là text của các ngôn ngữ và value là từ khóa của câu đó trong file properties languagesMap
    // từ từ khóa này thêm đuôi ngôn ngữ tương ứng sẽ hiển thị ra câu tương ứng bằng ngôn ngữ đó
    private final Map<String, String> languageMap = new HashMap<>();
    // đối tượng lưu cài đặt của chương trình
    private final Setup setup = new Setup();
    // đường dẫn file lưu cài đặt của chương trình
    private Path pathFile = Paths.get(FILE_SETUP_NAME);

    // lấy địa chỉ app data theo user người dùng và thêm vào thư mục Convert PDF to CHL
    // ví dụ: C:\Users\HuanTech PC\AppData\Roaming\convert pdf to chl
    private static final String appDataPath = System.getenv("APPDATA");
    // đường dẫn thư mục lưu cài đặt của chương trình từ địa chỉ app data đã lấy được ở trên
    private static final Path myAppPath = Paths.get(appDataPath, "Convert PDF to EXCEL");
    // list chứa các control UI cần để thay đổi ngôn ngữ hiển thị
    private final ObservableList<Object> controls = FXCollections.observableArrayList();

    /**
     * hàm khởi tạo đối tượng duy nhất của class
     */
    private SetupData() {
        // tạo đường dẫn đến file cài đặt từ thư mục lưu cài đặt
        FILE_SETUP_NAME = myAppPath.toAbsolutePath() + "\\setup_data.set";
        // tạo path của đường dẫn trên
        pathFile = Paths.get(FILE_SETUP_NAME);

        // tạo thư mục và file lưu cài đặt nếu chưa tồn tại
        createDirAndFile();

        // thêm các câu bằng 3 ngôn ngữ và từ khóa giống nhau của câu đó của nó vào map
        // 1 câu nhưng bằng 3 ngôn ngữ thì dùng từ khóa giống nhau để khi lấy từ khóa của nó rồi thêm đuôi ngôn ngữ(vi, ja, en) sẽ lấy được câu theo ngôn ngữ đó
        languageMap.put("Chọn file cần chuyển", "Select_the_file_to_transfer");
        languageMap.put("Chọn thư mục lưu file", "Select_the_folder_to_save_the_file");
        languageMap.put("THỰC HIỆN CHUYỂN FILE", "IMPLEMENT_FILE_TRANSFER");
        languageMap.put("Mở thư mục chứa các file CHL", "Open_the_folder_containing_CHL_files");
        languageMap.put("Danh sách file đã chuyển sang CHL.sysc2", "List_of_files_has_been_converted_to_chl_sysc");
        languageMap.put("Trợ giúp", "Help");
        languageMap.put("Chỉnh Sửa", "Edit");
        languageMap.put("Tệp", "File");
        languageMap.put("Xác nhận địa chỉ file PDF", "Confirm_the_PDF_address_file");
        languageMap.put("Địa chỉ của file PDF chưa được xác nhận", "The_address_of_the_PDF_file_has_not_been_confirmed");
        languageMap.put("Hãy chọn file PDF để tiếp tục!", "Please_select_the_PDF_file_to_continue");
        languageMap.put("Xác nhận thư mục chứa các file CHL", "Confirm_the_folder_containing_CHL_files");
        languageMap.put("Địa chỉ thư mục chứa các file CHL chưa được xác nhận", "Folder_address_containing_unconfirmed_CHL_files");
        languageMap.put("Hãy chọn thư mục chứa để tiếp tục!", "Please_select_the_containing_folder_to_continue");
        languageMap.put("Thông tin hoạt động chuyển file", "Information_on_file_transfer_activities");
        languageMap.put("Đã chuyển xong file PDF sang các file CHL", "Finished_converting_PDF_files_to_CHL_files");
        languageMap.put("Bạn có muốn mở thư mục chứa các file CHL và\ntự động copy địa chỉ không?", "Do_you_want_to_open_a_folder_containing_CHL_files_and_automatically_copy_the_address");
        languageMap.put("Thông báo lỗi chuyển file", "File_transfer_error_message");
        languageMap.put("Nội dung file PDF không phải là tính toán vật liệu hoặc file không được phép truy cập", "The_PDF_file_content_is_not_a_material_calculation_or_the_file_is_not_authorized_to_be_accessed");
        languageMap.put("Bạn có muốn chọn file khác và thực hiện lại không?", "Do_you_want_to_select_another_file_and_do_it_again");
        languageMap.put("CHUYỂN ĐỔI FILE PDF TÍNH TOÁN VẬT LIỆU SANG CHL", "Convert_material_calculation_PDF_files_to_CHL");
        languageMap.put("Lỗi mở thư mục", "Error_opening_folder");
        languageMap.put("Lỗi copy địa chỉ thư mục", "Error_copying_folder_address");
        languageMap.put("Thư mục chứa các file CHL có địa chỉ không đúng hoặc chưa được chọn!", "The_folder_containing_CHL_files_has_an_incorrect_address_or_has_not_been_selected");
        languageMap.put("Không thể copy địa chỉ thư mục chứa các file CHL", "Cannot_copy_folder_address_containing_CHL_files");
        languageMap.put("Copy link thư mục", "Copy_folder_link");
        languageMap.put("Đã copy link", "The_link_has_been_copied");
        languageMap.put("Giới thiệu", "About");
        languageMap.put("Đóng", "Close");
        languageMap.put("Giới thiệu:", "Introduce");
        languageMap.put("Phần mềm chuyển file PDF có nội dung tính vật liệu của thép hình sang các file CHL. Từ những thông tin lấy được trong file PDF các File CHL sẽ tạo định dạng phù hợp cho phần mềm CHL. Phần mềm CHL sẽ nhập file CHL vào để sử dụng.", "Software_to_convert_PDF");
        languageMap.put("Cách sử dụng:", "Using");
        languageMap.put("chọn địa chỉ file PDF có nội dung tính vật liệu trên máy và chọn địa chỉ thư mục sẽ chứa các file CHL khi chuyển xong. Các link này sau khi được chọn sẽ hiển thị ở các ô bên trái. Sau đó ấn vào nút chuyển để thực hiện. Các file CHL tạo ra sẽ hiển thị trong danh sách bên trái. Có thể nhấn nút Copy link thư mục sẽ chứa các file CHL hoặc ấn nút mở thư mục chứa các file CHL để mở cửa sổ thư mục này.", "Select_the_PDF");
        languageMap.put("Thực hiện: Lê Nhã", "copyright");
        languageMap.put("Tên file CHL đang tạo: (\"\") trùng tên với 1 file CHL khác đang được mở nên không thể ghi đè", "Name_of_the_CHL");
        languageMap.put("Hãy đóng file CHL đang mở để tiếp tục!", "Please_close_the_open_CHL_file_to_continue");
        languageMap.put("File CHL đang tạo: (\"\") có số dòng sản phẩm cần ghi lớn hơn 99 nên không thể ghi", "CHL_file_being_created");
        languageMap.put("Hãy chỉnh sửa lại dữ liệu vật liệu đang chuyển để tiếp tục!", "Please_edit_the_transferring_material_data_to_continue");

        languageMap.put("Tên file CHL", "CHL_file_name");
        languageMap.put("Sản phẩm(m)", "Product");
        languageMap.put("Vật liệu(m)", "Base_material");


        languageMap.put("転送するファイルを選択します", "Select_the_file_to_transfer");
        languageMap.put("ファイルを保存するフォルダーを選択します", "Select_the_folder_to_save_the_file");
        languageMap.put("ファイル転送を実行します", "IMPLEMENT_FILE_TRANSFER");
        languageMap.put("CHLファイルが入っているフォルダーを開きます", "Open_the_folder_containing_CHL_files");
        languageMap.put("ファイルのリストがCHL.sysc2に変換されました", "List_of_files_has_been_converted_to_chl_sysc");
        languageMap.put("ヘルプ", "Help");
        languageMap.put("編集", "Edit");
        languageMap.put("ファイル", "File");
        languageMap.put("PDFファイルの場所を確認", "Confirm_the_PDF_address_file");
        languageMap.put("PDFファイルのアドレスは未確認です", "The_address_of_the_PDF_file_has_not_been_confirmed");
        languageMap.put("続行するには PDF ファイルを選択してください。", "Please_select_the_PDF_file_to_continue");
        languageMap.put("CHLファイルが入っているフォルダを確認", "Confirm_the_folder_containing_CHL_files");
        languageMap.put("未確認のCHLファイルが入っているフォルダアドレス", "Folder_address_containing_unconfirmed_CHL_files");
        languageMap.put("続行するには、含まれているフォルダーを選択してください。", "Please_select_the_containing_folder_to_continue");
        languageMap.put("ファイル転送アクティビティに関する情報", "Information_on_file_transfer_activities");
        languageMap.put("PDFファイルからCHLファイルへの変換が完了しました", "Finished_converting_PDF_files_to_CHL_files");
        languageMap.put("CHL ファイルを含むフォルダーを開いて、アドレスを自動的にコピーしますか?", "Do_you_want_to_open_a_folder_containing_CHL_files_and_automatically_copy_the_address");
        languageMap.put("ファイル転送エラーメッセージ", "File_transfer_error_message");
        languageMap.put("PDFファイルの内容が材料計算ではないか、ファイルへのアクセスが許可されていません", "The_PDF_file_content_is_not_a_material_calculation_or_the_file_is_not_authorized_to_be_accessed");
        languageMap.put("別のファイルを選択してやり直しますか?", "Do_you_want_to_select_another_file_and_do_it_again");
        languageMap.put("材料計算書PDFファイルをCHLに変換", "Convert_material_calculation_PDF_files_to_CHL");
        languageMap.put("フォルダを開く際のエラー", "Error_opening_folder");
        languageMap.put("フォルダアドレスのコピーエラー", "Error_copying_folder_address");
        languageMap.put("CHL ファイルが含まれるフォルダーのアドレスが間違っているか、選択されていません。", "The_folder_containing_CHL_files_has_an_incorrect_address_or_has_not_been_selected");
        languageMap.put("CHLファイルを含むフォルダーアドレスをコピーできません", "Cannot_copy_folder_address_containing_CHL_files");
        languageMap.put("フォルダーリンクをコピー", "Copy_folder_link");
        languageMap.put("リンクがコピーされました", "The_link_has_been_copied");
        languageMap.put("情報", "About");
        languageMap.put("閉じる", "Close");
        languageMap.put("紹介します:", "Introduce");
        languageMap.put("形鋼の材料計算内容を記載したPDFファイルをCHLファイルに変換するソフトウェアです。 PDF ファイルで取得した情報に基づいて、CHL ファイルは CHL ソフトウェアに適した形式を作成します。 CHL ソフトウェアは CHL ファイルをインポートして使用します。", "Software_to_convert_PDF");
        languageMap.put("使用方法:", "Using");
        languageMap.put("パソコン上の材料計算コンテンツが含まれる PDF ファイルのアドレスを選択し、転送が完了したときに CHL ファイルが含まれるフォルダーのアドレスを選択します。これらのリンクを選択すると、左側のボックスに表示されます。その後、スイッチボタンを押して実行します。作成したCHLファイルが左側のリストに表示されます。 CHL ファイルを含むフォルダーへのリンクのコピー ボタンを押すか、CHL ファイルを含むフォルダーを開くボタンを押して、このフォルダー ウィンドウを開くことができます。", "Select_the_PDF");
        languageMap.put("作者: ル・ニャ", "copyright");
        languageMap.put("作成されているCHLファイルの名前: (\"\") は開いている別の CHL ファイルと同じ名前なので、上書きできません。", "Name_of_the_CHL");
        languageMap.put("続行するには、この開いている CHL ファイルを閉じてください。", "Please_close_the_open_CHL_file_to_continue");
        languageMap.put("CHL ファイルを作成中: (\"\") の製品ライン番号が 99 より大きいため、記録できません。", "CHL_file_being_created");
        languageMap.put("続行するには、転送中の鋼種データを編集してください。", "Please_edit_the_transferring_material_data_to_continue");

        languageMap.put("CHLファイル名", "CHL_file_name");
        languageMap.put("総製品(m)", "Product");
        languageMap.put("総鋼材(m)", "Base_material");


        languageMap.put("Select the file to transfer", "Select_the_file_to_transfer");
        languageMap.put("Select the folder to save the file", "Select_the_folder_to_save_the_file");
        languageMap.put("IMPLEMENT FILE TRANSFER", "IMPLEMENT_FILE_TRANSFER");
        languageMap.put("Open the folder containing CHL files", "Open_the_folder_containing_CHL_files");
        languageMap.put("List of files has been converted to CHL.sysc2", "List_of_files_has_been_converted_to_chl_sysc");
        languageMap.put("Help", "Help");
        languageMap.put("Edit", "Edit");
        languageMap.put("File", "File");
        languageMap.put("Confirm the PDF address file", "Confirm_the_PDF_address_file");
        languageMap.put("The address of the PDF file has not been confirmed", "The_address_of_the_PDF_file_has_not_been_confirmed");
        languageMap.put("Please select the PDF file to continue!", "Please_select_the_PDF_file_to_continue");
        languageMap.put("Confirm the folder containing CHL files", "Confirm_the_folder_containing_CHL_files");
        languageMap.put("Folder address containing unconfirmed CHL files", "Folder_address_containing_unconfirmed_CHL_files");
        languageMap.put("Please select the containing folder to continue!", "Please_select_the_containing_folder_to_continue");
        languageMap.put("Information on file transfer activities", "Information_on_file_transfer_activities");
        languageMap.put("Finished converting PDF files to CHL files", "Finished_converting_PDF_files_to_CHL_files");
        languageMap.put("Do you want to open a folder containing CHL files and automatically copy the address?", "Do_you_want_to_open_a_folder_containing_CHL_files_and_automatically_copy_the_address");
        languageMap.put("File transfer error message", "File_transfer_error_message");
        languageMap.put("The PDF file content is not a material calculation or the file is not authorized to be accessed", "The_PDF_file_content_is_not_a_material_calculation_or_the_file_is_not_authorized_to_be_accessed");
        languageMap.put("Do you want to select another file and do it again?", "Do_you_want_to_select_another_file_and_do_it_again");
        languageMap.put("CONVERT MATERIAL CALCULATION PDF FILES TO CHL", "Convert_material_calculation_PDF_files_to_CHL");
        languageMap.put("Error opening folder", "Error_opening_folder");
        languageMap.put("Error copying folder address", "Error_copying_folder_address");
        languageMap.put("The folder containing CHL files has an incorrect address or has not been selected!", "The_folder_containing_CHL_files_has_an_incorrect_address_or_has_not_been_selected");
        languageMap.put("Cannot copy folder address containing CHL files", "Cannot_copy_folder_address_containing_CHL_files");
        languageMap.put("Copy folder link", "Copy_folder_link");
        languageMap.put("The link has been copied", "The_link_has_been_copied");
        languageMap.put("About", "About");
        languageMap.put("Close", "Close");
        languageMap.put("Introduce:", "Introduce");
        languageMap.put("Software to convert PDF files containing material calculation content of shaped steel to CHL files. From the information obtained in the PDF file, the CHL File will create a suitable format for CHL software. CHL software will import the CHL file for use.", "Software_to_convert_PDF");
        languageMap.put("Using:", "Using");
        languageMap.put("Select the PDF file address containing the material calculation content on your computer and select the folder address that will contain the CHL files when the transfer is complete. These links, once selected, will be displayed in the boxes on the left. Then press the switch button to execute. The created CHL files will display in the list on the left. You can press the Copy link button to the folder that will contain CHL files or press the button to open the folder containing CHL files to open this folder window.", "Select_the_PDF");
        languageMap.put("copyright ©: Le Nha", "copyright");
        languageMap.put("Name of the CHL file being created: (\"\") has the same name as another CHL file that is currently open, so it cannot be overwritten", "Name_of_the_CHL");
        languageMap.put("Please close the open CHL file to continue!", "Please_close_the_open_CHL_file_to_continue");
        languageMap.put("CHL file being created: (\"\") has a product line number greater than 99, so it cannot be recorded", "CHL_file_being_created");
        languageMap.put("Please edit the transferring material data to continue!", "Please_edit_the_transferring_material_data_to_continue");

        languageMap.put("CHL file name", "CHL_file_name");
        languageMap.put("Product(m)", "Product");
        languageMap.put("Material(m)", "Base_material");
    }

    /**
     * @return đối tượng duy nhất(singleton) của SetupData
     */
    public static SetupData getInstance() {
        if (instance == null) {
            synchronized (SetupData.class) {
                instance = new SetupData();
            }
        }
        return instance;
    }

    /**
     * @return list các control
     */
    public ObservableList<Object> getControls() {
        return controls;
    }

    /**
     * @return đối tượng chứa cài đặt của app
     */
    public Setup getSetup() {
        return setup;
    }

    /**
     * set link của file pdf cho đối tượng cài đặt và lưu vào file
     * @param linkPdfFile link file pdf
     */
    public void setLinkPdfFile(String linkPdfFile) {
        setup.setLinkPdfFile(linkPdfFile);
        try {
            saveSetup();
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * set link thư mục chứa file chl sẽ tạo cho đối tượng cài đặt và lưu vào file
     * @param SaveCvsFileDir link thư mục chứa file chl sẽ tạo
     */
    public void setLinkSaveCvsFileDir(String SaveCvsFileDir) {
        setup.setLinkSaveCvsFileDir(SaveCvsFileDir);
        try {
            saveSetup();
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * set ngôn ngữ cho đối tượng cài đặt và lưu ngôn ngữ vào file
     * @param lang ngôn ngữ
     */
    public void setLang(String lang) throws IOException {
        setup.setLang(lang);
        saveSetup();
    }

    /**
     * @return list chứa các file chl đã tạo
     */
    public ObservableList<ExcelFile> getExcelFile() {
        return excelFiles;
    }

    /**
     * @return map ngôn ngữ
     */
    public Map<String, String> getLanguageMap() {
        return languageMap;
    }

    /**
     * lấy các cài đặt từ file cài đặt và lưu vào đối tượng cài đặt
     * @throws IOException lỗi đọc file cài đặt
     */
    public void loadSetup() throws IOException {

        // Tạo thư mục và file nếu nó không tồn tại
        if (Files.notExists(pathFile)) {
            createDirAndFile();
            System.out.println(setup.getLinkPdfFile() + setup.getLinkSaveCvsFileDir() + setup.getLang());
            return;
        }

        // đọc dữ liệu nhị phân từ file cài đặt
        try (DataInputStream dis = new DataInputStream(new BufferedInputStream(Files.newInputStream(pathFile)))) {
            boolean eof = false;
            while (!eof) {
                try {

                    String linkPdfFile = dis.readUTF();
                    String linkSaveCvsFileDir = dis.readUTF();
                    String lang = dis.readUTF();

                    setup.setLinkPdfFile(linkPdfFile);
                    System.out.println(":" + setup.getLinkPdfFile());

                    setup.setLinkSaveCvsFileDir(linkSaveCvsFileDir);
                    System.out.println(":" + setup.getLinkSaveCvsFileDir());

                    setup.setLang(lang);
                    System.out.println(":" + setup.getLang());

                } catch (EOFException e) {
                    eof = true;
                }
            }
        }
    }

    /**
     * lưu dữ liệu cài đặt từ đối tượng setup vào file
     * @throws IOException lỗi ghi file
     */
    public void saveSetup() throws IOException {
        // Tạo thư mục và file nếu nó không tồn tại
        createDirAndFile();

        // Tạo đối tượng File đại diện cho file cần xóa
        File file = new File(FILE_SETUP_NAME);

        // Kiểm tra nếu file tồn tại thì xóa nó
        // vì file là readonly nên cần xóa đi tạo file mới
        if (file.exists()) {
            if (file.delete()) {
                System.out.println("File data đã được xóa thành công.");
            } else {
                System.out.println("Xóa file data thất bại.");
            }
        } else {
            System.out.println("File data không tồn tại.");
        }
        // ghi file nhị phân
        try (DataOutputStream dos = new DataOutputStream(new BufferedOutputStream(Files.newOutputStream(pathFile)))) {
            dos.writeUTF(setup.getLinkPdfFile());
            dos.writeUTF(setup.getLinkSaveCvsFileDir());
            dos.writeUTF(setup.getLang());
        }

        // Đặt quyền chỉ đọc cho file, để không chỉnh sửa file
        File readOnly = new File(FILE_SETUP_NAME);
        if (readOnly.exists()) {
            boolean result = readOnly.setReadOnly();
            if (result) {
                System.out.println("File data is set to read-only.");
            } else {
                System.out.println("Failed to set file data to read-only.");
            }
        } else {
            System.out.println("File data does not exist.");
        }
    }

    /**
     * tạo thư mục và file nếu nó không tồn tại
     */
    private void createDirAndFile() {
        try {
            // Tạo thư mục nếu nó không tồn tại
            if (Files.notExists(myAppPath)) {
                Files.createDirectory(myAppPath);
            }
            // Tạo file nếu nó không tồn tại
            if (Files.notExists(pathFile)) {
                Files.createFile(pathFile);
                System.out.println(pathFile);
            }

        } catch (IOException e) {
            System.out.println(e.getMessage());
            e.printStackTrace();
        }

    }


}

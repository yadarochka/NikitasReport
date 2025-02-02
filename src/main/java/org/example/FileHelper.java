package org.example;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;

public class FileHelper {

    public enum FILE_TYPE {
        INPUT(ConfigLoader.getInputFilePath()),
        OUTPUT(ConfigLoader.getOutputFilePath());

        private final String filePath;

        FILE_TYPE(String filePath) {
            this.filePath = filePath;
        }

        public String getFilePath() {
            return filePath;
        }
    }
    private static final String inputFileURL = ConfigLoader.getInputFileURL();

    private static final Logger logger = LogManager.getLogger(FileHelper.class);

    public static void downloadFile() {

        File file = new File(inputFileURL);

        if (file.exists()) {
            logger.debug("Файл с расписанием уже скачан");
        } else {
            logger.debug("Скачивание файла расписания...");
            try {
                URL url = new URL(inputFileURL);
                HttpURLConnection connection = getHttpURLConnection(url);
                connection.disconnect();
                logger.debug("Файл успешно скачан: {}", FILE_TYPE.INPUT.filePath);
            } catch (IOException e) {
                logger.fatal("Ошибка скачивания: {}", e.getMessage());
            }
        }
    }

    private static HttpURLConnection getHttpURLConnection(URL url) throws IOException {
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("GET");

        try (InputStream inputStream = connection.getInputStream();
             FileOutputStream outputStream = new FileOutputStream(FILE_TYPE.INPUT.filePath)) {
            byte[] buffer = new byte[1024];
            int bytesRead;
            while ((bytesRead = inputStream.read(buffer)) != -1) {
                outputStream.write(buffer, 0, bytesRead);
            }
        }
        return connection;
    }

    public static void deleteFile(FILE_TYPE fileType) {

        try {
            Files.delete(Path.of(fileType.getFilePath()));
        } catch (NoSuchFileException e) {
            System.out.println("Файл не найден.");
        } catch (IOException e) {
            System.out.println("Ошибка при удалении файла: " + e.getMessage());
        }

        logger.info("Файл {} успешно удалён.", fileType);
    }

    public static void openOutputFile(){
        String batchFilePath = "cmd.bat"; // Укажи путь к .bat

        ProcessBuilder processBuilder = new ProcessBuilder(batchFilePath);
        processBuilder.directory(new File(System.getProperty("user.dir"))); // Устанавливаем рабочую директорию
        processBuilder.redirectErrorStream(true); // Перенаправляем ошибки в стандартный вывод

        try {
            Process process = processBuilder.start();
            process.waitFor(); // Ожидаем завершения выполнения
            System.out.println("Batch script executed successfully!");
        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
        }
    }
}
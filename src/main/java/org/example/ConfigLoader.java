package org.example;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;
import java.io.File;

public class ConfigLoader {
    private static Properties properties = new Properties();
    private static String ConfigPath = "src/main/java/org/example/config.properties";
    private static final Logger logger = LogManager.getLogger(ConfigLoader.class);

    static {
        File file = new File(ConfigPath);

        if (file.exists()) {
            logger.debug("Файл конфига найден!");
        } else {
            logger.fatal("Файл конфига НЕ найден!");
            throw new RuntimeException("Файл конфига НЕ найден!");
        }
        try (FileInputStream input = new FileInputStream(ConfigPath)) {
            properties.load(input);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static String getInputFileURL() {
        return properties.getProperty("inputFileURL");
    }

    public static String getInputFilePath() {
        return properties.getProperty("inputFilePath");
    }

    public static String getTemplateFilePath() {
        return properties.getProperty("templateFilePath");
    }

    public static String getOutputFilePath() {
        return properties.getProperty("outputFilePath");
    }
}
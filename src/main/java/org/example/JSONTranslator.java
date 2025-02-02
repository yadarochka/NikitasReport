package org.example;

import com.google.gson.Gson;

import java.io.FileReader;
import java.io.IOException;
import java.util.Map;

public class JSONTranslator {
    private static Map<String, String> translations;

    static {
        Gson gson = new Gson();
        try {
            String filePath = "groupName.JSON";
            translations = gson.fromJson(new FileReader(filePath), Map.class);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static String translate(String word, String defaultValue) {
        return translations.getOrDefault(word, defaultValue);
    }

}


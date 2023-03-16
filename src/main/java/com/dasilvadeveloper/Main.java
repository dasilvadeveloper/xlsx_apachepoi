package com.dasilvadeveloper;

import com.dasilvadeveloper.model.Person;
import com.dasilvadeveloper.util.XlsxFactory;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.IOException;
import java.util.ArrayList;

public class Main {
    public static void main(String[] args) throws JsonProcessingException {

        ObjectMapper mapper = new ObjectMapper();
        ArrayList<Person> persons = mapper.readValue(personsJson, new TypeReference<ArrayList<Person>>() {
        });

        System.out.println(persons);

        // create xlsx factory
        XlsxFactory xlsxFactory = new XlsxFactory(XlsxFactory.Layout.SIMPLE);

        // set headers to display
        // IMPORTANT note for setHeaderColumnAttributes
        // . <<<< each attribute must exactly match the attribute of the object >>>>
        xlsxFactory.setHeaderColumnAttributes(new String[]{"id", "name", "age"});

        //set header titles
        xlsxFactory.setHeaderTitles(new String[]{"ID", "Name", "Age"});

        // set data model
        xlsxFactory.setDataModel(persons);

        try{
            // generate & save xls
            xlsxFactory.build("/files/xls/", "New Employees", true);
        } catch (IOException | NoSuchFieldException | IllegalAccessException e) {
            System.err.println(e.getMessage());
        }

    }

    static String personsJson = "[\n" +
            "  {\n" +
            "    \"id\": \"641084ceb6415085e5d28184\",\n" +
            "    \"name\": \"Bright Blair\",\n" +
            "    \"age\": 34\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084ce7d88d31ea47f4d70\",\n" +
            "    \"name\": \"Glover Snider\",\n" +
            "    \"age\": 33\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084ceb86dd369e94a1e6b\",\n" +
            "    \"name\": \"Brenda Rodgers\",\n" +
            "    \"age\": 39\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084ce2d292bd46d618d8c\",\n" +
            "    \"name\": \"Bonner Buckley\",\n" +
            "    \"age\": 28\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084ce72d3174bf70aa2d7\",\n" +
            "    \"name\": \"Susanna Humphrey\",\n" +
            "    \"age\": 40\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084cea7d6fecc30d33112\",\n" +
            "    \"name\": \"Gladys Medina\",\n" +
            "    \"age\": 22\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084ce1372cbe5d6be844d\",\n" +
            "    \"name\": \"Church Reed\",\n" +
            "    \"age\": 38\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084ceb099930def1ab96c\",\n" +
            "    \"name\": \"Ines Guy\",\n" +
            "    \"age\": 25\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084ce5b2ca1bb99909cad\",\n" +
            "    \"name\": \"Nicholson Austin\",\n" +
            "    \"age\": 20\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084cee405c728f79fe655\",\n" +
            "    \"name\": \"Pace Mcgee\",\n" +
            "    \"age\": 33\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084ce299152a8313408e5\",\n" +
            "    \"name\": \"Solomon Stark\",\n" +
            "    \"age\": 39\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084ce5adf41ce28304d58\",\n" +
            "    \"name\": \"Polly Forbes\",\n" +
            "    \"age\": 34\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084ce3e0f346554e69284\",\n" +
            "    \"name\": \"Cheri Hendrix\",\n" +
            "    \"age\": 40\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084ce692482a6e640be2f\",\n" +
            "    \"name\": \"Pickett Bruce\",\n" +
            "    \"age\": 37\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084cee6e2521a12ebf9fd\",\n" +
            "    \"name\": \"Lizzie Suarez\",\n" +
            "    \"age\": 40\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084ce629b37d7267fbd5f\",\n" +
            "    \"name\": \"Ortiz Short\",\n" +
            "    \"age\": 35\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084ce0d51dfc3f2a3646c\",\n" +
            "    \"name\": \"Augusta Peterson\",\n" +
            "    \"age\": 37\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084ceab4b5293d3033ba4\",\n" +
            "    \"name\": \"Harriett Pugh\",\n" +
            "    \"age\": 40\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084cedbe005a0e61b6432\",\n" +
            "    \"name\": \"Maricela Willis\",\n" +
            "    \"age\": 39\n" +
            "  },\n" +
            "  {\n" +
            "    \"id\": \"641084cef265bf4420b421d3\",\n" +
            "    \"name\": \"Vicky Cook\",\n" +
            "    \"age\": 36\n" +
            "  }\n" +
            "]";
}

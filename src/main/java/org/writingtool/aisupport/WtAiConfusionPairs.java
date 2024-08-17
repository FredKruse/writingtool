/* WritingTool, a LibreOffice Extension based on LanguageTool
 * Copyright (C) 2024 Fred Kruse (https://fk-es.de)
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301
 * USA
 */
package org.writingtool.aisupport;

import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.languagetool.JLanguageTool;
import org.writingtool.tools.WtMessageHandler;

import com.sun.star.lang.Locale;

public class WtAiConfusionPairs {

  
  private static String getConfusionSetFilePath(Locale locale, String fileName) {
    return "/" + locale.Language + "/" + fileName;
  }
  
  /**
   * get the list of words out of spelling.txt files defined by LT
   */
  public static Map<String, Set<String>> getConfusionWordMap(Locale locale, String fileName) {
    return getConfusionWordMap(locale, fileName, null);
  }

  public static Map<String, Set<String>> getConfusionWordMap(Locale locale, String fileName, Map<String, Set<String>> words) {
    if (words == null) {
      words = new HashMap<>();
    }
    String path = getConfusionSetFilePath(locale, fileName);
    WtMessageHandler.printToLogFile("Confusion set path: " + path);
    if (JLanguageTool.getDataBroker().resourceExists(path)) {
      List<String> lines = JLanguageTool.getDataBroker().getFromResourceDirAsLines(path);
      if (lines != null) {
        for (String line : lines) {
          line = line.trim();
          if (!line.isEmpty()) {
            if (line.startsWith("#")) {
              line = line.substring(1);
            }
            String[] lineString = line.split("#");
            String wLine; 
            if (lineString.length == 2) {
              wLine = lineString[0].trim();
            } else {
              continue;
            }
            String[] sWords = wLine.split(";");
            String word1;
            String word2;
            if (sWords.length == 3) {
              word1 = sWords[0].trim();
              word2 = sWords[1].trim();
            } else if (sWords.length == 2) {
              String[] tmpWords = sWords[0].trim().split("->");
              word1 = tmpWords[0].trim();
              word2 = tmpWords[1].trim();
            } else {
              continue;
            }
            if (!word1.equalsIgnoreCase(word2)) {
              if (words.containsKey(word1)) {
                Set<String> wList = words.get(word1);
                wList.add(word2);
              } else if (words.containsKey(word2)) {
                Set<String> wList = words.get(word2);
                wList.add(word1);
              } else {
                Set<String> wList = new HashSet<String>();
                wList.add(word2);
                words.put(word1, wList);
              }
              WtMessageHandler.printToLogFile("Word pair added: " + word1 + ", " + word2);
            }
          }
        }
      }
      
    } else {
      WtMessageHandler.printToLogFile("Confusion path doe not exist: " + path);
    }
    return words;
  }

}

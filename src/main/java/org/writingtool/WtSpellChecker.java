/* LanguageTool, a natural language style checker
 * Copyright (C) 2011 Daniel Naber (http://www.danielnaber.de)
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
package org.writingtool;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.languagetool.JLanguageTool;
import org.languagetool.JLanguageTool.ParagraphHandling;
import org.languagetool.Language;
import org.languagetool.rules.Rule;
import org.languagetool.rules.RuleMatch;
import org.languagetool.rules.spelling.SpellingCheckRule;
import org.languagetool.rules.spelling.hunspell.HunspellRule;
import org.languagetool.rules.spelling.morfologik.MorfologikSpellerRule;
import org.writingtool.CacheIO.SpellCache;
import org.writingtool.OfficeTools.OfficeProductInfo;
import org.writingtool.config.Configuration;

import com.sun.star.beans.PropertyValue;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XServiceDisplayName;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.lib.uno.helper.Factory;
import com.sun.star.lib.uno.helper.WeakBase;
import com.sun.star.linguistic2.XSpellAlternatives;
import com.sun.star.linguistic2.XSpellChecker;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.uno.XComponentContext;

/**
 * Class for spell checking by LanguageTool
 * @since 6.1
 * @author Fred Kruse
 */
public class WtSpellChecker extends WeakBase implements XServiceInfo, 
  XServiceDisplayName, XSpellChecker {

  private static boolean DEBUG_MODE = false;  // set to true for debug output
  
  private static final int MAX_WRONG = 10000;
  private static final String PROB_CHARS = ".*[~<>].*";
  
  // Service name required by the OOo API && our own name.
  private static final String[] SERVICE_NAMES = {
          "com.sun.star.linguistic2.SpellChecker",
          OfficeTools.WT_SPELL_SERVICE_NAME };
  
  private static JLanguageTool lt = null;
  private static Locale lastLocale = null;                //  locale for spell check
  private static SpellingCheckRule spellingCheckRule = null;
  private static MorfologikSpellerRule mSpellRule = null;
  private static HunspellRule hSpellRule = null;
  private static final Map<String, List<String>> lastWrongWords = new HashMap<>();
  private static final Map<String, List<String[]>> lastSuggestions = new HashMap<>();
  private static String last1 = new String();
  private static String last2 = new String();
  private static XComponentContext xContext = null;
  private static boolean noLtSpeller = false;
  
  public WtSpellChecker(XComponentContext xContxt) {
    if (xContext == null) {
      try {
        xContext = xContxt;
        MessageHandler.init(xContext);
        OfficeProductInfo officeInfo = OfficeTools.getOfficeProductInfo(xContext);
        if (officeInfo == null || officeInfo.osArch.equals("x86")
            || !isEnoughHeap() || !runLTSpellChecker(xContext)) {
          noLtSpeller = true;
        } else {
          CacheIO c = new CacheIO(); 
          SpellCache sc = c.new SpellCache();
          if (sc.read()) {
            if (sc.getWrongWords() != null && sc.getSuggestions() != null
                && sc.getWrongWords().size() == sc.getSuggestions().size()) {
              for (String loc : sc.getWrongWords().keySet()) {
                if (!sc.getWrongWords().get(loc).isEmpty()) {
                  List<String> savedLastWords = sc.getWrongWords().get(loc);
                  List<String[]> savedSuggestions = sc.getSuggestions().get(loc);
                  List<String> lastWords = new ArrayList<>();
                  List<String[]> suggestions = new ArrayList<>();
                  for (int i = 0; i < sc.getWrongWords().get(loc).size() && i < MAX_WRONG; i++) {
                    lastWords.add(savedLastWords.get(i));
                    suggestions.add(savedSuggestions.get(i));
                  }
                  lastWrongWords.put(loc, lastWords);
                  lastSuggestions.put(loc, suggestions);
                }
              }
            }
          }
        }
      } catch (Throwable e) {
        MessageHandler.showError(e);
      }
    }
  }

  /**
   * Get XSingleComponentFactory
   * Default method called by LO/OO extensions
   */
  public static XSingleComponentFactory __getComponentFactory(String sImplName) {
    SingletonFactory xFactory = null;
    if (sImplName.equals(WtSpellChecker.class.getName())) {
      xFactory = new SingletonFactory(true);
    }
    return xFactory;
  }

  /**
   * Write keys to registry
   * Default method called by LO/OO extensions
   */
  public static boolean __writeRegistryServiceInfo(XRegistryKey regKey) {
    return Factory.writeRegistryServiceInfo(WtSpellChecker.class.getName(), WtSpellChecker.getServiceNames(), regKey);
  }

  @Override
  public Locale[] getLocales() {
    if (noLtSpeller) {
      return new Locale[0];
    }
    return MultiDocumentsHandler.getLocales();
  }

  @Override
  public boolean hasLocale(Locale locale) {
    return MultiDocumentsHandler.hasLocale(locale);
  }

  @Override
  public String getImplementationName() {
    return WtSpellChecker.class.getName();
  }

  /**
   * Get the names of supported services
   * interface: XServiceInfo
   */
  @Override
  public String[] getSupportedServiceNames() {
    return getServiceNames();
  }

  /**
   * Get the LT service names
   */
  static String[] getServiceNames() {
    return SERVICE_NAMES;
  }

  /**
   * Test if the service is supported by LT
   * interface: XServiceInfo
   */
  @Override
  public boolean supportsService(String sServiceName) {
    for (String sName : SERVICE_NAMES) {
      if (sServiceName.equals(sName)) {
        return true;
      }
    }
    return false;
  }

  /**
   * is a correct spelled word
   */
  @Override
  public boolean isValid(String word, Locale locale, PropertyValue[] Properties) throws IllegalArgumentException {
    try {
      if (noLtSpeller || locale == null || !hasLocale(locale)) {
        return false;
      }
      if (word == null || word.trim().isEmpty()) {
        if (DEBUG_MODE) {
          MessageHandler.printToLogFile("LtSpellChecker: isValid: word id empty: " + (word == null ? "null" : word));
        }
        return true;
      }
      if (DEBUG_MODE) {
        MessageHandler.printToLogFile("LtSpellChecker: isValid: test word/string: '" + (word == null ? "null" : word) + "'");
      }
      if (word.matches(PROB_CHARS)) {
        if (DEBUG_MODE) {
          MessageHandler.printToLogFile("LtSpellChecker: isValid: Problematic word found: " + (word == null ? "null" : word));
        }
        return false;
      }
      String[] words = word.split(" ");
      if (words.length == 2 && words[0].equals(words[1])) {  //  this a workaround for a possible problem in LO
        if (last2.endsWith(".") || words[0].equals(last1)) {
          word = words[0] + " " + last2;
          if (DEBUG_MODE) {
            MessageHandler.printToLogFile("LtSpellChecker: isValid: repeated word changed to: " + (word == null ? "null" : word));
          }
        } else {
          if (DEBUG_MODE) {
            MessageHandler.printToLogFile("LtSpellChecker: isValid: repeated word found: " + (word == null ? "null" : word));
          }
          return false;
        }
      }
      last1 = last2;
      last2 = word;
      String localeStr = OfficeTools.localeToString(locale);
      List<String> wrongWords = lastWrongWords.get(localeStr);
      if (wrongWords != null && wrongWords.contains(word)) {
        if (DEBUG_MODE) {
          MessageHandler.printToLogFile("LtSpellChecker: isValid: invalid word found in list: " + (word == null ? "null" : word));
        }
        return false;
      }
      initSpellChecker(locale);
      if (spellingCheckRule != null) {
        if (words.length == 1 && !spellingCheckRule.isMisspelled(word)) {
          if (DEBUG_MODE) {
            MessageHandler.printToLogFile("LtSpellChecker: isValid: valid word found: " + (word == null ? "null" : word));
          }
          return true;
        }
        List<RuleMatch> matches = lt.check(word,true, ParagraphHandling.ONLYNONPARA);
        if (matches == null || matches.size() == 0) {
          if (DEBUG_MODE) {
            MessageHandler.printToLogFile("LtSpellChecker: isValid: valid word found (matches == 0): " + (word == null ? "null" : word));
          }
          return true;
        }
        if (wrongWords == null) {
          lastWrongWords.put(localeStr, new ArrayList<String>());
          lastSuggestions.put(localeStr, new ArrayList<String[]>());
        }
        if (!lastWrongWords.get(localeStr).contains(word)) {
          lastWrongWords.get(localeStr).add(new String(word));
          lastSuggestions.get(localeStr).add(suggestionsToArray(matches.get(0).getSuggestedReplacements()));
          if (lastWrongWords.get(localeStr).size() >= MAX_WRONG) {
            lastWrongWords.get(localeStr).remove(0);
            lastSuggestions.get(localeStr).remove(0);
          }
        }
        if (DEBUG_MODE) {
          MessageHandler.printToLogFile("LtSpellChecker: isValid: invalid word found: " + (word == null ? "null" : word));
        }
        return false;
      }
    } catch (Throwable e) {
      MessageHandler.showError(e);
    }
    MessageHandler.printToLogFile("LtSpellChecker: isValid: Problem with spelling rule: word: " + (word == null ? "null" : word));
    return false;
  }

  /**
   * get alternatives for a non correct spelled word
   */
  @Override
  public XSpellAlternatives spell(String word, Locale locale, PropertyValue[] properties) throws IllegalArgumentException {
    LTSpellAlternatives alternatives = new LTSpellAlternatives(word, locale);
    return alternatives;
  }

  /**
   * Init spell checker for locale
   */
  private void initSpellChecker(Locale locale) {
    try {
      if (lastLocale == null || !OfficeTools.isEqualLocale(lastLocale, locale)) {
        lastLocale = locale;
        Language lang = MultiDocumentsHandler.getLanguage(locale);
        lt = new JLanguageTool(lang);
        for (Rule rule : lt.getAllRules()) {
          if (rule.isDictionaryBasedSpellingRule()) {
            spellingCheckRule = (SpellingCheckRule) rule;
            if (spellingCheckRule instanceof MorfologikSpellerRule) {
              mSpellRule = (MorfologikSpellerRule) spellingCheckRule;
              hSpellRule = null;
            } else if (spellingCheckRule instanceof HunspellRule) {
              hSpellRule = (HunspellRule) spellingCheckRule;
              mSpellRule = null;
            }
          } else {
            lt.disableRule(rule.getId());
          }
        }
      }
    } catch (Throwable e) {
      MessageHandler.showError(e);
    }
  }
  
  /**
   * Convert list of suggestions to array and reduce it size
   */
  private String[] suggestionsToArray(List<String> suggestions) {
    int numSuggestions = suggestions.size();
    String[] allSuggestions = suggestions.toArray(new String[numSuggestions]);
    if (allSuggestions.length > OfficeTools.MAX_SUGGESTIONS) {
      allSuggestions = Arrays.copyOfRange(allSuggestions, 0, OfficeTools.MAX_SUGGESTIONS);
    }
    return allSuggestions;
  }
  
  /**
   * class for getting spelling alternatives
   */
  class LTSpellAlternatives implements XSpellAlternatives {
    
    Locale locale;
    String word;
    String[] alternatives = null;
    
    LTSpellAlternatives(String word, Locale locale) {
      try {
        this.word = word;
        this.locale = locale;
        if (noLtSpeller) {
          alternatives = new String[0];
          return;
        }
        String localeStr = OfficeTools.localeToString(locale);
        if (word == null || word.trim().isEmpty() || localeStr == null || localeStr.isEmpty()) {
          alternatives = new String[0];
          return;
        }
        List<String> wrongWords = lastWrongWords.get(localeStr);
        List<String[]> suggestions = lastSuggestions.get(localeStr);
        if (wrongWords != null && suggestions != null) {
          if (wrongWords.contains(word)) {
            int num = wrongWords.indexOf(word);
            alternatives = suggestions.get(num);
            if (alternatives == null) {
              alternatives = new String[0];
            }
            return;
          }
        }
        if (word.matches(PROB_CHARS)) {
          if (DEBUG_MODE) {
            MessageHandler.printToLogFile("LtSpellChecker: LTSpellAlternatives: Problematic word found: " + (word == null ? "null" : word));
          }
          alternatives = new String[0];
          return;
        }
        if (mSpellRule != null) {
          alternatives = mSpellRule.getSpellingSuggestions(word).toArray(new String[0]);
        }
        if (hSpellRule != null) {
          alternatives = hSpellRule.getSuggestions(word).toArray(new String[0]);
        }
      } catch (Throwable t) {
        MessageHandler.showError(t);
      }
      if (alternatives == null) {
        alternatives = new String[0];
      }
    }

    @Override
    public String[] getAlternatives() {
      return alternatives;
    }

    @Override
    public short getAlternativesCount() {
      return (short) alternatives.length;
    }

    @Override
    public short getFailureType() {
      // ??? unclear
      return 0;
    }

    @Override
    public Locale getLocale() {
      if (locale == null) {
        return lastLocale;
      }
      return locale;
    }

    @Override
    public String getWord() {
      if (word == null) {
        return "";
      }
      return word;
    }
    
  }

  @Override
  public String getServiceDisplayName(Locale locale) {
    return MultiDocumentsHandler.getServiceDisplayName(locale);
  }
  
  public static Map<String, List<String>> getWrongWords() {
    return lastWrongWords;
  }
   
  public static Map<String, List<String[]>> getSuggestions() {
    return lastSuggestions;
  }
  
  public static boolean runLTSpellChecker(XComponentContext xContext) {
    Configuration confg;
    try {
      confg = new Configuration(OfficeTools.getLOConfigDir(xContext), 
          OfficeTools.CONFIG_FILE, OfficeTools.getOldConfigFile(), null, true);
      OfficeTools.setLogLevel(confg.getlogLevel());
      DEBUG_MODE = OfficeTools.DEBUG_MODE_SP;
      return confg.useLtSpellChecker();
    } catch (IOException e) {
      MessageHandler.printToLogFile("Can't read configuration: LT spell checker not used!");
    }
    return false;
  }

  public static boolean isEnoughHeap() {
    int maxHeapSpace = (int) (OfficeTools.getMaxHeapSpace()/1048576);
    boolean ret = maxHeapSpace >= OfficeTools.SPELL_CHECK_MIN_HEAP;
    if (!ret) {
      MessageHandler.printToLogFile("Heap Space (" + maxHeapSpace + ") is too small: LT spell checker not used!\n"
          + "Set heap space greater than " +  OfficeTools.SPELL_CHECK_MIN_HEAP);
    }
    return ret;
  }

  public static void resetSpellCache() {
    lastWrongWords.clear();
    lastSuggestions.clear();
  }
   
}

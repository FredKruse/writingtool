/* LanguageTool, a natural language style checker
 * Copyright (C) 2014 Daniel Naber (http://www.danielnaber.de)
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

import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;

import org.languagetool.AnalyzedSentence;
import org.languagetool.rules.RuleMatch;
import org.writingtool.DocumentCache;
import org.writingtool.LinguisticServices;
import org.writingtool.MessageHandler;
import org.writingtool.OfficeTools;
import org.writingtool.ResultCache;
import org.writingtool.SingleCheck;
import org.writingtool.SingleDocument;
import org.writingtool.SwJLanguageTool;
import org.writingtool.ViewCursorTools;
import org.writingtool.config.Configuration;

import com.sun.star.lang.Locale;
import com.sun.star.linguistic2.SingleProofreadingError;

/**
 * Class to detect errors by a AI API
 * @since 6.5
 * @author Fred Kruse
 */
public class AiErrorDetection {
  
  boolean debugModeTm = true;
  boolean debugMode = OfficeTools.DEBUG_MODE_AI;
  
  private static final ResourceBundle messages = OfficeTools.getMessageBundle();
  private final SingleDocument document;
  private final DocumentCache docCache;
  private final Configuration config;
  private final SwJLanguageTool lt;
//  private final int minParaLength = (int) (AiRemote.CORRECT_INSTRUCTION.length() * 1.2);
  private static String lastLanguage = null;
  private static String correctCommand = null;
  
  public AiErrorDetection(SingleDocument document, Configuration config, SwJLanguageTool lt) {
    this.document = document;
    this.config = config;
    this.lt = lt;
    docCache = document.getDocumentCache();
  }
  
  public void addAiRuleMatchesForParagraph() {
    if (docCache != null) {
      if (debugMode) {
        MessageHandler.printToLogFile("AiErrorDetection: addAiRuleMatchesForParagraph: start");
      }
      ViewCursorTools viewCursor = new ViewCursorTools(document.getXComponent());
      int nFPara = docCache.getFlatParagraphNumber(viewCursor.getViewCursorParagraph());
      addAiRuleMatchesForParagraph(nFPara);
    } else {
      MessageHandler.printToLogFile("AiErrorDetection: addAiRuleMatchesForParagraph: docCache == null");
    }
  }

  public void addAiRuleMatchesForParagraph(int nFPara) {
    if (docCache == null || nFPara < 0) {
      return;
    }
    try {
      String paraText = docCache.getFlatParagraph(nFPara);
      int[] footnotePos = docCache.getFlatParagraphFootnotes(nFPara);
      List<Integer> deletedChars = docCache.getFlatParagraphDeletedCharacters(nFPara);
      if (paraText == null || paraText.trim().isEmpty()) {
        if (debugMode) {
          MessageHandler.printToLogFile("AiErrorDetection: addAiRuleMatchesForParagraph: nFPara " + nFPara + " is empty: return");
        }
        addMatchesByAiRule(nFPara, null, footnotePos, deletedChars);
        return;
      }
      Locale locale = docCache.getFlatParagraphLocale(nFPara);
      if (lastLanguage == null || !lastLanguage.equals(locale.Language)) {
        lastLanguage = new String(locale.Language);
        correctCommand = AiRemote.getInstruction(AiRemote.CORRECT_INSTRUCTION, locale);
//        MessageHandler.printToLogFile("AiErrorDetection: addAiRuleMatchesForParagraph: correctCommand: " + correctCommand);
      }
      RuleMatch[] ruleMatches = getAiRuleMatchesForParagraph(nFPara, paraText, locale, footnotePos, deletedChars);
      if (debugMode && ruleMatches != null) {
        MessageHandler.printToLogFile("AiErrorDetection: addAiRuleMatchesForParagraph: nFPara: " + nFPara + ", rulematches: " + ruleMatches.length);
      }
      addMatchesByAiRule(nFPara, ruleMatches, footnotePos, deletedChars);
    } catch (Throwable t) {
      MessageHandler.showError(t);
    }
  }

  public void addAiRuleMatchesForParagraph(String paraText, Locale locale, int[] footnotePos, List<Integer> deletedChars) {
    try {
      RuleMatch[] ruleMatches = getAiRuleMatchesForParagraph(-1, paraText, locale, footnotePos, deletedChars);
      addMatchesByAiRule(-1, ruleMatches, footnotePos, deletedChars);
    } catch (Throwable t) {
      MessageHandler.showError(t);
    }
  }
  
  public List<RuleMatch> getListAiRuleMatchesForParagraph(int nFPara, String paraText, 
      Locale locale, int[] footnotePos, List<Integer> deletedChars) throws Throwable {
    List<RuleMatch> matchList = new ArrayList<>();
    RuleMatch[] ruleMatches = getAiRuleMatchesForParagraph(nFPara, paraText, locale, footnotePos, deletedChars);
    if (ruleMatches != null) {
      for (RuleMatch match : ruleMatches) {
        matchList.add(match);
      }
    }
    return matchList;
  }
    
  public RuleMatch[] getAiRuleMatchesForParagraph(int nFPara, String paraText, 
      Locale locale, int[] footnotePos, List<Integer> deletedChars) throws Throwable {
    if (docCache == null) {
      return null;
    }
    if (paraText == null || paraText.trim().isEmpty()) {
      if (debugMode) {
        MessageHandler.printToLogFile("AiErrorDetection: getAiRuleMatchesForParagraph: paraText: " + (paraText == null? "NULL" : "EMPTY"));
      }
      return null;
    }
    paraText = DocumentCache.fixLinebreak(SingleCheck.removeFootnotes(paraText, 
        footnotePos, deletedChars));
    List<AnalyzedSentence> analyzedSentences;
    if (nFPara < 0) {
      paraText = DocumentCache.fixLinebreak(SingleCheck.removeFootnotes(paraText, 
          footnotePos, deletedChars));
      analyzedSentences =  lt.analyzeText(paraText.replace("\u00AD", ""));
    } else {
      analyzedSentences = docCache.getAnalyzedParagraph(nFPara);
    }
    if (analyzedSentences == null) {
      analyzedSentences = docCache.createAnalyzedParagraph(nFPara, lt);
      if (analyzedSentences == null) {
        if (debugMode) {
          MessageHandler.printToLogFile("AiErrorDetection: getAiRuleMatchesForParagraph: analyzedSentences == null");
        }
        return null;
      }
    }
    return getMatchesByAiRule(nFPara, paraText, analyzedSentences, locale, footnotePos, deletedChars);
  }
    
  private RuleMatch[] getMatchesByAiRule(int nFPara, String paraText, List<AnalyzedSentence> analyzedSentences,
      Locale locale, int[] footnotePos, List<Integer> deletedChars) throws Throwable {
    String result = getAiResult(paraText, locale);
    if (result == null || result.trim().isEmpty()) {
      if (debugMode) {
        MessageHandler.printToLogFile("AiErrorDetection: getMatchesByAiRule: result: " + (result == null? "NULL" : "EMPTY"));
      }
      return null;
    }
    if (debugMode) {
      MessageHandler.printToLogFile("\nAiErrorDetection: getMatchesByAiRule: result: " + result + "\n");
    }
    long startTime = 0;
    if (debugModeTm) {
      startTime = System.currentTimeMillis();
    }
    List<AnalyzedSentence> analyzedAiResult =  lt.analyzeText(result.replace("\u00AD", ""));
    AiDetectionRule aiRule = getAiDetectionRule(result, analyzedAiResult, 
        document.getMultiDocumentsHandler().getLinguisticServices(), locale , messages, config.aiShowStylisticChanges());
    RuleMatch[] matches = aiRule.match(analyzedSentences);
    if (debugModeTm) {
      long runTime = System.currentTimeMillis() - startTime;
      MessageHandler.printToLogFile("AiErrorDetection: getMatchesByAiRule: Time to run AI detection rule: " + runTime);
    }
    if (debugMode) {
      MessageHandler.printToLogFile("AiErrorDetection: getMatchesByAiRule: matches: " + matches.length);
    }
    return matches;
  }
    
  private void addMatchesByAiRule(int nFPara, RuleMatch[] ruleMatches,
                    int[] footnotePos, List<Integer> deletedChars) throws Throwable {
    ResultCache aiCache = document.getParagraphsCache().get(OfficeTools.CACHE_AI);
    if (ruleMatches == null || ruleMatches.length == 0) {
      aiCache.put(nFPara, null, new SingleProofreadingError[0]);
      return;
    }
    List<SingleProofreadingError> errorList = new ArrayList<>();
    for (RuleMatch myRuleMatch : ruleMatches) {
      if (debugMode) {
        MessageHandler.printToLogFile("Rule match suggestion: " + myRuleMatch.getSuggestedReplacements().get(0));
      }
      SingleProofreadingError error = SingleCheck.createOOoError(myRuleMatch, 0, footnotePos, null, config);
      if (debugMode) {
        MessageHandler.printToLogFile("error suggestion: " + error.aSuggestions[0]);
      }
      errorList.add(SingleCheck.correctRuleMatchWithFootnotes(
          error, footnotePos, deletedChars));
    }
    aiCache.put(nFPara, null, errorList.toArray(new SingleProofreadingError[0]));
    List<Integer> changedParas = new ArrayList<>();
    changedParas.add(nFPara);
    document.remarkChangedParagraphs(changedParas, changedParas, false);
  }
    
  private String getAiResult(String para, Locale locale) throws Throwable {
    if (para == null || para.isEmpty()) {
      return "";
    }
//    String text = CORRECT_COMMAND + ": " + para;
//    MessageHandler.showMessage("Input is: " + text);
    AiRemote aiRemote = new AiRemote(config);
    String output = aiRemote.runInstruction(correctCommand, para, locale, true);
//    String output = aiRemote.runInstruction(AiRemote.CORRECT_INSTRUCTION, para, locale, true);
    return output;
  }
/*  
  private void translateCorrectCommand(Locale locale) throws Throwable {
    if (lastLanguage == null || !lastLanguage.equals(locale.Language)) {
      lastLanguage = new String(locale.Language);
      if (lastLanguage.equals("en")) {
        correctCommand = AiRemote.CORRECT_COMMAND;
      } else {
        Language lang = MultiDocumentsHandler.getLanguage(locale);
        String languageName = lang.getName();
        String command = AiRemote.TRANSLATE_COMMAND + languageName;
        MessageHandler.printToLogFile("AiErrorDetection: translateCorrectCommand: command: " + command);
        AiRemote aiRemote = new AiRemote(config);
        correctCommand = aiRemote.runInstruction(command, AiRemote.CORRECT_COMMAND, true);
        if (correctCommand.endsWith(".")) {
          correctCommand = correctCommand.substring(0, correctCommand.length() - 1);
        }
        MessageHandler.printToLogFile("AiErrorDetection: translateCorrectCommand: correctCommand: " + correctCommand);
      }
    }
  }
*/
  
  private AiDetectionRule getAiDetectionRule(String aiResultText, List<AnalyzedSentence> analyzedAiResult, 
      LinguisticServices linguServices, Locale locale, ResourceBundle messages, boolean showStylisticHints) {
    try {
      Class<?>[] cArgs = { String.class, List.class, LinguisticServices.class, Locale.class, ResourceBundle.class, boolean.class };
      Class<?> clazz = Class.forName("org.writingtool.aisupport.AiDetectionRule_" + locale.Language);
      MessageHandler.printToLogFile("Use detection rule for: " + locale.Language);
      return (AiDetectionRule) clazz.getDeclaredConstructor(cArgs).newInstance(aiResultText, 
          analyzedAiResult, linguServices, locale, messages, showStylisticHints);
    } catch (Exception e) {
      MessageHandler.printException(e);
      MessageHandler.printToLogFile("Use general detection rule");
      return new AiDetectionRule(aiResultText, analyzedAiResult, linguServices, locale, messages, showStylisticHints);
    }
    
  }
  
  
  
}

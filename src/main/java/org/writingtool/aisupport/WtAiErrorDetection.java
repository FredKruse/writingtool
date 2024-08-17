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

import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;

import org.languagetool.AnalyzedSentence;
import org.languagetool.AnalyzedTokenReadings;
import org.languagetool.JLanguageTool.ParagraphHandling;
import org.languagetool.rules.RuleMatch;
import org.writingtool.WtDocumentCache;
import org.writingtool.WtLinguisticServices;
import org.writingtool.WtResultCache;
import org.writingtool.WtSingleCheck;
import org.writingtool.WtSingleDocument;
import org.writingtool.WtResultCache.CacheEntry;
import org.writingtool.WtLanguageTool;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtOfficeTools.RemoteCheck;
import org.writingtool.tools.WtViewCursorTools;

import com.sun.star.lang.Locale;
import com.sun.star.linguistic2.SingleProofreadingError;

/**
 * Class to detect errors by a AI API
 * @since 6.5
 * @author Fred Kruse
 */
public class WtAiErrorDetection {
  
  private static final int MIN_WORD = 4;
  
  private boolean debugModeTm = WtOfficeTools.DEBUG_MODE_TM;
  private boolean debugMode = WtOfficeTools.DEBUG_MODE_AI;
  
  private static final ResourceBundle messages = WtOfficeTools.getMessageBundle();
  private final WtSingleDocument document;
  private final WtDocumentCache docCache;
  private final WtConfiguration config;
  private final WtLanguageTool lt;
//  private final int minParaLength = (int) (AiRemote.CORRECT_INSTRUCTION.length() * 1.2);
  private static String lastLanguage = null;
  private static String correctCommand = null;
  
  public WtAiErrorDetection(WtSingleDocument document, WtConfiguration config, WtLanguageTool lt) {
    this.document = document;
    this.config = config;
    this.lt = lt;
    docCache = document.getDocumentCache();
  }
  
  public void addAiRuleMatchesForParagraph() {
    if (docCache != null) {
      if (debugMode) {
        WtMessageHandler.printToLogFile("AiErrorDetection: addAiRuleMatchesForParagraph: start");
      }
      WtViewCursorTools viewCursor = new WtViewCursorTools(document.getXComponent());
      int nFPara = docCache.getFlatParagraphNumber(viewCursor.getViewCursorParagraph());
      addAiRuleMatchesForParagraph(nFPara);
    } else {
      WtMessageHandler.printToLogFile("AiErrorDetection: addAiRuleMatchesForParagraph: docCache == null");
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
          WtMessageHandler.printToLogFile("AiErrorDetection: addAiRuleMatchesForParagraph: nFPara " + nFPara + " is empty: return");
        }
        addMatchesByAiRule(nFPara, null, footnotePos, deletedChars);
        return;
      }
      Locale locale = docCache.getFlatParagraphLocale(nFPara);
      if (lastLanguage == null || !lastLanguage.equals(locale.Language)) {
        lastLanguage = new String(locale.Language);
        correctCommand = WtAiRemote.getInstruction(WtAiRemote.CORRECT_INSTRUCTION, locale);
//        MessageHandler.printToLogFile("AiErrorDetection: addAiRuleMatchesForParagraph: correctCommand: " + correctCommand);
      }
      RuleMatch[] ruleMatches = getAiRuleMatchesForParagraph(nFPara, paraText, locale, footnotePos, deletedChars);
      if (debugMode && ruleMatches != null) {
        WtMessageHandler.printToLogFile("AiErrorDetection: addAiRuleMatchesForParagraph: nFPara: " + nFPara + ", rulematches: " + ruleMatches.length);
      }
      addMatchesByAiRule(nFPara, ruleMatches, footnotePos, deletedChars);
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }

  public void addAiRuleMatchesForParagraph(String paraText, Locale locale, int[] footnotePos, List<Integer> deletedChars) {
    try {
      RuleMatch[] ruleMatches = getAiRuleMatchesForParagraph(-1, paraText, locale, footnotePos, deletedChars);
      addMatchesByAiRule(-1, ruleMatches, footnotePos, deletedChars);
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
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
        WtMessageHandler.printToLogFile("AiErrorDetection: getAiRuleMatchesForParagraph: paraText: " + (paraText == null? "NULL" : "EMPTY"));
      }
      return null;
    }
    paraText = WtDocumentCache.fixLinebreak(WtSingleCheck.removeFootnotes(paraText, 
        footnotePos, deletedChars));
    List<AnalyzedSentence> analyzedSentences;
    if (nFPara < 0) {
      paraText = WtDocumentCache.fixLinebreak(WtSingleCheck.removeFootnotes(paraText, 
          footnotePos, deletedChars));
      analyzedSentences =  lt.analyzeText(paraText.replace("\u00AD", ""));
    } else {
      analyzedSentences = docCache.getAnalyzedParagraph(nFPara);
    }
    if (analyzedSentences == null) {
      analyzedSentences = docCache.createAnalyzedParagraph(nFPara, lt);
      if (analyzedSentences == null) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("AiErrorDetection: getAiRuleMatchesForParagraph: analyzedSentences == null");
        }
        return null;
      }
    }
    // Don't analyze a paragraph with less than MIN_WORD words
    if (analyzedSentences.size() < 2) {
      if (analyzedSentences.size() == 0) {
        return new RuleMatch[0];
      }
      int n = 0;
      for (AnalyzedTokenReadings token : analyzedSentences.get(0).getTokensWithoutWhitespace()) {
        if (!token.isNonWord()) {
          n++;
        }
      }
      if (n <= MIN_WORD) {
        return new RuleMatch[0];
      }
    }
    // ---
    return getMatchesByAiRule(nFPara, paraText, analyzedSentences, locale, footnotePos, deletedChars);
  }
    
  private RuleMatch[] getMatchesByAiRule(int nFPara, String paraText, List<AnalyzedSentence> analyzedSentences,
      Locale locale, int[] footnotePos, List<Integer> deletedChars) throws Throwable {
    String result = getAiResult(paraText, locale);
    if (result == null || result.trim().isEmpty()) {
      if (debugMode) {
        WtMessageHandler.printToLogFile("AiErrorDetection: getMatchesByAiRule: result: " + (result == null? "NULL" : "EMPTY"));
      }
      return null;
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("\nAiErrorDetection: getMatchesByAiRule: result: " + result + "\n");
    }
    long startTime = 0;
    if (debugModeTm) {
      startTime = System.currentTimeMillis();
    }
    List<AnalyzedSentence> analyzedAiResult =  lt.analyzeText(result.replace("\u00AD", ""));
    WtAiDetectionRule aiRule = getAiDetectionRule(result, analyzedAiResult, 
        document.getMultiDocumentsHandler().getLinguisticServices(), locale , messages, config.aiShowStylisticChanges());
    RuleMatch[] matches = aiRule.match(analyzedSentences);
    matches = filterRuleMatches(matches, result, locale, analyzedAiResult);
    if (debugModeTm) {
      long runTime = System.currentTimeMillis() - startTime;
      WtMessageHandler.printToLogFile("AiErrorDetection: getMatchesByAiRule: Time to run AI detection rule: " + runTime);
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("AiErrorDetection: getMatchesByAiRule: matches: " + matches.length);
    }
    return matches;
  }
    
  private void addMatchesByAiRule(int nFPara, RuleMatch[] ruleMatches,
                    int[] footnotePos, List<Integer> deletedChars) throws Throwable {
    WtResultCache aiCache = document.getParagraphsCache().get(WtOfficeTools.CACHE_AI);
    CacheEntry cEntry = aiCache.getCacheEntry(nFPara);
    boolean isMatch = cEntry != null && cEntry.errorArray.length > 0;
    if (ruleMatches == null || ruleMatches.length == 0) {
      aiCache.put(nFPara, null, new SingleProofreadingError[0]);
    } else {
      List<SingleProofreadingError> errorList = new ArrayList<>();
      for (RuleMatch myRuleMatch : ruleMatches) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("Rule match suggestion: " + myRuleMatch.getSuggestedReplacements().get(0));
        }
        SingleProofreadingError error = WtSingleCheck.createOOoError(myRuleMatch, 0, footnotePos, null, config);
        if (debugMode) {
          WtMessageHandler.printToLogFile("error suggestion: " + error.aSuggestions[0]);
        }
        errorList.add(WtSingleCheck.correctRuleMatchWithFootnotes(
            error, footnotePos, deletedChars));
      }
      aiCache.put(nFPara, null, errorList.toArray(new SingleProofreadingError[0]));
    }
    if (isMatch) {
      List<Integer> changedParas = new ArrayList<>();
      changedParas.add(nFPara);
      document.remarkChangedParagraphs(changedParas, changedParas, false);
    }
  }
    
  private String getAiResult(String para, Locale locale) throws Throwable {
    if (para == null || para.isEmpty()) {
      return "";
    }
    WtAiRemote aiRemote = new WtAiRemote(document.getMultiDocumentsHandler(),config);
    String output = aiRemote.runInstruction(correctCommand, para, 0, 1, locale, true);
    return output;
  }
  
  private boolean isCorrectResult(RuleMatch match, List<RuleMatch> resultMatches) {
    if (resultMatches != null) {
      for (RuleMatch rMatch : resultMatches) {
        if (
            (rMatch.getFromPos() >= match.getFromPos() && rMatch.getFromPos() < match.getToPos()) ||
            (rMatch.getToPos() >= match.getFromPos() && rMatch.getToPos() <= match.getToPos()) ||
            (match.getFromPos() >= rMatch.getFromPos() && match.getFromPos() < rMatch.getToPos()) ||
            (match.getToPos() >= rMatch.getFromPos() && match.getToPos() <= rMatch.getToPos())
            ) {
//          if (debugMode) {
            WtMessageHandler.printToLogFile("WtErrorDetection: filterRuleMatches: incorrect match: suggestion: " +
              match.getSuggestedReplacements().get(0) + ", reason: " + rMatch.getMessage());
//            }
          return false;
        }
      }
    }
    return true;
  }
  
  private RuleMatch[] filterRuleMatches(RuleMatch[] matches, String result, Locale locale,
      List<AnalyzedSentence> analyzedAiResult) throws Throwable {
    if (matches == null || matches.length == 0) {
      return matches;
    }
    List<RuleMatch> resultMatches = lt.check(result, analyzedAiResult, ParagraphHandling.NORMAL, RemoteCheck.ALL);
    if (resultMatches == null) {
      return matches;
    }
    for (int i = resultMatches.size() - 1; i >= 0; i--) {
      if (resultMatches.get(i).getRule().isDictionaryBasedSpellingRule()) {
        WtLinguisticServices linguServices = document.getMultiDocumentsHandler().getLinguisticServices();
        String word = result.substring(resultMatches.get(i).getFromPos(), resultMatches.get(i).getToPos());
        if(linguServices.isCorrectSpell(word, locale)) {
          resultMatches.remove(i);
        }
      }
    }
    if (resultMatches.size() == 0) {
      return matches;
    }
    List<RuleMatch> correctMatches = new ArrayList<>();
    for (RuleMatch match : matches) {
      if (isCorrectResult(match, resultMatches)) {
        correctMatches.add(match);
      }
    }
    return correctMatches.toArray(new RuleMatch[correctMatches.size()]);
  }
  
  private WtAiDetectionRule getAiDetectionRule(String aiResultText, List<AnalyzedSentence> analyzedAiResult, 
      WtLinguisticServices linguServices, Locale locale, ResourceBundle messages, boolean showStylisticHints) {
    try {
      Class<?>[] cArgs = { String.class, List.class, WtLinguisticServices.class, Locale.class, ResourceBundle.class, boolean.class };
      Class<?> clazz = Class.forName("org.writingtool.aisupport.WtAiDetectionRule_" + locale.Language);
//      WtMessageHandler.printToLogFile("Use detection rule for: " + locale.Language);
      return (WtAiDetectionRule) clazz.getDeclaredConstructor(cArgs).newInstance(aiResultText, 
          analyzedAiResult, linguServices, locale, messages, showStylisticHints);
    } catch (Exception e) {
      WtMessageHandler.printException(e);
      WtMessageHandler.printToLogFile("Use general detection rule");
      return new WtAiDetectionRule(aiResultText, analyzedAiResult, linguServices, locale, messages, showStylisticHints);
    }
    
  }
  
  
  
}

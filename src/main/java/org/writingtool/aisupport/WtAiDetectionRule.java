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

import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.ResourceBundle;
import java.util.regex.Pattern;

import org.languagetool.AnalyzedSentence;
import org.languagetool.AnalyzedToken;
import org.languagetool.AnalyzedTokenReadings;
import org.languagetool.Tag;
import org.languagetool.rules.Categories;
import org.languagetool.rules.ITSIssueType;
import org.languagetool.rules.RuleMatch;
import org.languagetool.rules.RuleMatch.Type;
import org.writingtool.WtLinguisticServices;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.languagetool.rules.TextLevelRule;

import com.sun.star.lang.Locale;

/**
 * Rule to detect errors by a AI API
 * @since 6.5
 * @author Fred Kruse
 */
public class WtAiDetectionRule extends TextLevelRule {

  public static final int STYLE_HINT_LIMIT = 25;  //  Limit of changed tokens in a sentence in percent (after that a style hint is assumed)
  
  private static final Pattern QUOTES = Pattern.compile("[\"“”“„»«]");
  private static final Pattern SINGLE_QUOTES = Pattern.compile("[‚‘’'›‹]");
  private static final Pattern PUNCTUATION = Pattern.compile("[,.!?:]");
  private static final Pattern OPENING_BRACKETS = Pattern.compile("[{(\\[]");

  private boolean debugMode = WtOfficeTools.DEBUG_MODE_AI;   //  should be false except for testing

  private final ResourceBundle messages;
  private final String aiResultText;
  private final List<AnalyzedSentence> analyzedAiResult;
  private final String ruleMessage;
  private final boolean showStylisticHints;
  private final WtLinguisticServices linguServices;
  private final Locale locale;

  
  WtAiDetectionRule(String aiResultText, List<AnalyzedSentence> analyzedAiResult, 
      WtLinguisticServices linguServices, Locale locale, ResourceBundle messages, boolean showStylisticHints) {
    this.aiResultText = aiResultText;
    this.analyzedAiResult = analyzedAiResult;
    this.messages = messages;
    this.showStylisticHints = showStylisticHints;
    this.linguServices = linguServices;
    this.locale = locale;
    ruleMessage = messages.getString("loAiRuleMessage");
    
    setCategory(Categories.STYLE.getCategory(messages));
    setLocQualityIssueType(ITSIssueType.Grammar);
    setTags(Collections.singletonList(Tag.picky));
    
    if (debugMode) {
      WtMessageHandler.printToLogFile("AiDetectionRule: showStylisticHints: " + showStylisticHints);
    }
  }

  /**
   * Override this function for specific Languages
   * zzz -> no specific language
   */
  public String getLanguage() {
    return "zzz";
  }
  
  private boolean isIgnoredToken(String paraToken, String resultToken) {
/*
    if (resultToken.equals("\"")) {
      return QUOTES.matcher(paraToken).matches();
    }
    if (resultToken.equals("'")) {
      return SINGLE_QUOTES.matcher(paraToken).matches();
    }
*//*
    if (resultToken.equals("\"") || resultToken.equals("'")) {
      return QUOTES.matcher(paraToken).matches() || SINGLE_QUOTES.matcher(paraToken).matches();
    }
*/
    if (QUOTES.matcher(resultToken).matches() || SINGLE_QUOTES.matcher(resultToken).matches()) {
      return QUOTES.matcher(paraToken).matches() || SINGLE_QUOTES.matcher(paraToken).matches();
    }
    if (resultToken.equals("-")) {
      return paraToken.equals("–");
    }
    return false;
  }

  public boolean isQuote(String token) {
    return (QUOTES.matcher(token).matches() || SINGLE_QUOTES.matcher(token).matches());
  }

  @Override
  public RuleMatch[] match(List<AnalyzedSentence> sentences) throws IOException {
    List<RuleMatch> matches = new ArrayList<>();
    List<AiRuleMatch> tmpMatches = new ArrayList<>();
    List<WtAiToken> paraTokens = new ArrayList<>();
    List<Integer> sentenceEnds = new ArrayList<>();
    int nSenTokens = 0;
    int nSentence = 0;
//    int lastSentenceStart = 0;
    int lastResultStart = 0;
    int pos = 0;
    int sEnd = 0;
    int sugStart = 0;
    int sugEnd = 0;
    boolean mergeSentences = false;
    for (AnalyzedSentence sentence : sentences) {
      AnalyzedTokenReadings[] tokens = sentence.getTokensWithoutWhitespace();
      for (int i = 1; i < tokens.length; i++) {
        paraTokens.add(new WtAiToken(tokens[i], pos, sentence));
        sEnd++;
      }
      pos += sentence.getCorrectedTextLength();
      sentenceEnds.add(sEnd);
    }
    List<WtAiToken> resultTokens = new ArrayList<>();
    pos = 0;
    for (AnalyzedSentence sentence : analyzedAiResult) {
      AnalyzedTokenReadings[] tokens = sentence.getTokensWithoutWhitespace();
      for (int i = 1; i < tokens.length; i++) {
        resultTokens.add(new WtAiToken(tokens[i], pos, sentence));
      }
      pos += sentence.getCorrectedTextLength();
    }
    int i;
    int j = 0;
    for (i = 0; i < paraTokens.size() && j < resultTokens.size(); i++) {
      if (!paraTokens.get(i).getToken().equals(resultTokens.get(j).getToken()) 
          && !isIgnoredToken(paraTokens.get(i).getToken(), resultTokens.get(j).getToken())) {
        if ((i == 0 && ("{".equals(resultTokens.get(j).getToken()) || QUOTES.matcher(resultTokens.get(j).getToken()).matches()))
            && j + 1 < resultTokens.size() && paraTokens.get(i).getToken().equals(resultTokens.get(j + 1).getToken())) {
          j += 2;
          continue;
        }
        if (isQuote(paraTokens.get(i).getToken())
            && i + 1 < paraTokens.size() && paraTokens.get(i + 1).getToken().equals(resultTokens.get(j).getToken())) {
          continue;
        }
        int posStart = paraTokens.get(i).getStartPos();
        int nParaTokenStart = i;
        int nParaTokenEnd = 0;
        int nResultTokenStart = 0;
        int nResultTokenEnd = 0;
        AnalyzedSentence sentence = paraTokens.get(i).getSentence();
        int posEnd = 0;
        String suggestion = null;
        WtAiToken singleWordToken = null;
        boolean endFound = false;
        for (int n = 1; !endFound && i + n < paraTokens.size() && j + n < resultTokens.size(); n++) {
          for (int i1 = i + n; !endFound && i1 >= i; i1--) {
            for(int j1 = j + n; j1 >= j; j1--) {
              if (paraTokens.get(i1).getToken().equals(resultTokens.get(j1).getToken()) 
                  || isIgnoredToken(paraTokens.get(i1).getToken(), resultTokens.get(j1).getToken())) {
                endFound = true;
                if (i1 - 1 < i) {
                  if (i > 0) {
                    posStart = paraTokens.get(i - 1).getStartPos();
                    posEnd = paraTokens.get(i1 - 1).getEndPos();
                    sugStart = resultTokens.get(j - 1).getStartPos();
                    sugEnd = resultTokens.get(j1 - 1).getEndPos();
                    singleWordToken = j == j1 ? resultTokens.get(j1 - 1) : null;
                    nParaTokenStart = i - 1;
                    nParaTokenEnd = i1 - 1;
                    nResultTokenStart = j - 1;
                    nResultTokenEnd = j1 -1;
                  } else {
                    posEnd = paraTokens.get(i1).getEndPos();
                    nParaTokenEnd = i1;
                    if (j < 1) {
                      j = 1;
                    }
                    if (j1 < 0) {
                      j1 = 0;
                    }
                    sugStart = resultTokens.get(j - 1).getStartPos();
                    sugEnd = resultTokens.get(j1).getEndPos();
                    nResultTokenStart = j - 1;
                    nResultTokenEnd = j1;
                    singleWordToken = j - 1 == j1 ? resultTokens.get(j1) : null;
                  }
                } else {
                  posEnd = paraTokens.get(i1 - 1).getEndPos();
                  nParaTokenEnd = i1 - 1;
                  if (j < 0) {
                    j = 0;
                  }
                  if (j1 < 1) {
                    j1 = 1;
                  }
                  if (j <= j1 - 1) {
                    sugStart = resultTokens.get(j).getStartPos();
                    sugEnd = resultTokens.get(j1 - 1).getEndPos();
                    nResultTokenStart = j;
                    nResultTokenEnd = j1 -1;
                    singleWordToken = j == j1 - 1 ? resultTokens.get(j) : null;
                  } else {
                    if (i > 0 && !PUNCTUATION.matcher(paraTokens.get(i - 1).getToken()).matches()) {
                      posStart = paraTokens.get(i - 1).getEndPos();
                    } else {
                      posEnd = paraTokens.get(i1).getStartPos();
                    }
                    sugStart = resultTokens.get(j1).getEndPos();
                    sugEnd = resultTokens.get(j1).getEndPos();
                    nResultTokenStart = j1;
                    nResultTokenEnd = j1 -1;
                  }
                }
                nSenTokens += (i1 - i + 1);
                j = j1;
                i = i1;
                break;
              }
            }
          }
        }
        if (!endFound) {
          posEnd = paraTokens.get(paraTokens.size() - 1).getEndPos();
          nParaTokenEnd = paraTokens.size() - 1;
          sugStart = resultTokens.get(j).getStartPos();
          sugEnd = resultTokens.get(resultTokens.size() - 1).getEndPos();
          nResultTokenStart = j;
          nResultTokenEnd = resultTokens.size() - 1;
          j = resultTokens.size() - 1;
        }
        suggestion = sugStart >= sugEnd ? "" : aiResultText.substring(sugStart, sugEnd);
        if (!isEndQoute(j, resultTokens) && isCorrectSuggestion(suggestion, singleWordToken)
            && !isMatchException(nParaTokenStart, nParaTokenEnd, nResultTokenStart, nResultTokenEnd, paraTokens, resultTokens)) {
          RuleMatch ruleMatch = new RuleMatch(this, sentence, posStart, posEnd, ruleMessage);
          ruleMatch.addSuggestedReplacement(suggestion);
          ruleMatch.setType(Type.Hint);
          setType(nParaTokenStart, nParaTokenEnd, nResultTokenStart, nResultTokenEnd, paraTokens, resultTokens, ruleMatch);
          tmpMatches.add(new AiRuleMatch(ruleMatch, sugStart, sugEnd, nParaTokenStart, nParaTokenEnd, nResultTokenStart, nResultTokenEnd));
          if (debugMode) {
            WtMessageHandler.printToLogFile("AiDetectionRule: match: found: start: " + posStart + ", end: " + posEnd
                + ", suggestion: " + suggestion);
          }
        } else if (debugMode) {
          WtMessageHandler.printToLogFile("AiDetectionRule: match: not correct spell: locale: " + WtOfficeTools.localeToString(locale)
              + ", suggestion: " + suggestion);
        }
      }
      j++;
      if (i >= sentenceEnds.get(nSentence)) {
        mergeSentences = true;
        while (i >= sentenceEnds.get(nSentence)) {
          nSentence++;
        }
      }
      if (i == sentenceEnds.get(nSentence) - 1) {
        if (nSenTokens > 0) {
          int allSenTokens = nSentence == 0 ? sentenceEnds.get(nSentence) : sentenceEnds.get(nSentence) - sentenceEnds.get(nSentence - 1);
          if (mergeSentences || styleHintAssumed(nSenTokens, allSenTokens, tmpMatches, paraTokens, resultTokens)) {
            if (showStylisticHints && !tmpMatches.isEmpty()) {
              int startPos = tmpMatches.get(0).ruleMatch.getFromPos();
              int endPos = tmpMatches.get(tmpMatches.size() - 1).ruleMatch.getToPos();
              RuleMatch ruleMatch = new RuleMatch(this, null, startPos, endPos, ruleMessage);
              int suggestionStart = tmpMatches.get(0).suggestionStart;
              int suggestionEnd = tmpMatches.get(tmpMatches.size() - 1).suggestionEnd;
              String suggestion = aiResultText.substring(suggestionStart, suggestionEnd);
              ruleMatch.addSuggestedReplacement(suggestion);
              ruleMatch.setType(Type.Other);
              matches.add(ruleMatch);
              if (debugMode) {
                WtMessageHandler.printToLogFile("AiDetectionRule: match: Stylistic hint: suggestion: " + suggestion);
              }
            }
            mergeSentences = false;
          } else {
            addAllRuleMatches(matches, tmpMatches);
            if (debugMode) {
              WtMessageHandler.printToLogFile("AiDetectionRule: match: add matches: " + tmpMatches.size()
              + ", total: " + matches.size());
            }
          }
        }
        tmpMatches.clear();
        nSenTokens = 0;
        nSentence++;
//        lastSentenceStart = i + 1;
        lastResultStart = j;
      }
    }
    if (tmpMatches.size() > 0 || (j < resultTokens.size() 
            && (!paraTokens.get(i - 1).getToken().equals(resultTokens.get(j - 1).getToken())
                || (resultTokens.get(j).isNonWord() && !PUNCTUATION.matcher(resultTokens.get(j).getToken()).matches())
        ))) {
//        || (!"}".equals(resultTokens.get(j).getToken()) && !QUOTES.matcher(resultTokens.get(j).getToken()).matches() 
//            && !OPENING_BRACKETS.matcher(resultTokens.get(j).getToken()).matches())))) {
      nSenTokens++;
      if (nSentence > 0) {
        nSentence--;
      }
      boolean overSentenceEnd = j < resultTokens.size() && PUNCTUATION.matcher(resultTokens.get(j - 1).getToken()).matches();
      int allSenTokens = nSentence == 0 ? sentenceEnds.get(nSentence) : sentenceEnds.get(nSentence) - sentenceEnds.get(nSentence - 1);
      if (debugMode) {
        WtMessageHandler.printToLogFile("AiDetectionRule: match: j < resultTokens.size(): mergeSentences: " + mergeSentences
            + ", nSenTokens: " + nSenTokens + ", allSenTokens: " + allSenTokens);
      }
      if (mergeSentences || overSentenceEnd || styleHintAssumed(nSenTokens, allSenTokens, tmpMatches, paraTokens, resultTokens)) {
        if (showStylisticHints) {
          int startPos;
          int endPos;
          int suggestionStart;
          int suggestionEnd;
          if (tmpMatches.size() > 0) {
            startPos = tmpMatches.get(0).ruleMatch.getFromPos();
            endPos = tmpMatches.get(tmpMatches.size() - 1).ruleMatch.getToPos();
            suggestionStart = tmpMatches.get(0).suggestionStart;
            suggestionEnd = tmpMatches.get(tmpMatches.size() - 1).suggestionEnd;
          } else {
            startPos = nSentence == 0 ? paraTokens.get(0).getStartPos() : paraTokens.get(sentenceEnds.get(nSentence - 1)).getStartPos();
            endPos = paraTokens.get(paraTokens.size() - 1).getEndPos();
            suggestionStart = resultTokens.get(lastResultStart).getStartPos();
            suggestionEnd = resultTokens.get(resultTokens.size() - 1).getEndPos();
          }
          RuleMatch ruleMatch = new RuleMatch(this, null, startPos, endPos, ruleMessage);
          String suggestion = aiResultText.substring(suggestionStart, suggestionEnd);
          ruleMatch.addSuggestedReplacement(suggestion);
          ruleMatch.setType(Type.Other);
          matches.add(ruleMatch);
          if (debugMode) {
            WtMessageHandler.printToLogFile("AiDetectionRule: match: Stylistic hint: suggestion: " + suggestion);
          }
        }
      } else {
        int j1;
        for (j1 = j + 1; j1 < resultTokens.size() && !OPENING_BRACKETS.matcher(resultTokens.get(j1).getToken()).matches(); j1++);
        if (j1 > resultTokens.size()) {
          j1 = resultTokens.size();
        }
        String suggestion = aiResultText.substring(resultTokens.get(j - 1).getStartPos(), resultTokens.get(j1 - 1).getEndPos());
        if (suggestion.isEmpty() || j != j1 || resultTokens.get(j - 1).isNonWord() || linguServices.isCorrectSpell(suggestion, locale)) {
          RuleMatch ruleMatch = new RuleMatch(this, null, paraTokens.get(paraTokens.size() - 1).getStartPos(), 
              paraTokens.get(paraTokens.size() - 1).getEndPos(), ruleMessage);
          ruleMatch.addSuggestedReplacement(suggestion);
          setType(paraTokens.size() - 1, paraTokens.size() - 1, j - 1, j1 - 1, paraTokens, resultTokens, ruleMatch);
          tmpMatches.add(new AiRuleMatch(ruleMatch, resultTokens.get(j - 1).getStartPos(), resultTokens.get(resultTokens.size() - 1).getEndPos(),
              paraTokens.size() - 1, paraTokens.size() - 1, j - 1, j1 - 1));
        }
        addAllRuleMatches(matches, tmpMatches);
        if (debugMode) {
          WtMessageHandler.printToLogFile("AiDetectionRule: match: add matches: " + tmpMatches.size()
          + ", total: " + matches.size());
        }
      }
    }
    if (debugMode && j < resultTokens.size()) {
      WtMessageHandler.printToLogFile("AiDetectionRule: match: j < resultTokens.size(): paraTokens.get(i - 1): " + paraTokens.get(i - 1).getToken()
          + ", resultTokens.get(j - 1): " + resultTokens.get(j - 1).getToken());
    }
    return toRuleMatchArray(matches);
  }
/*  
  private AnalyzedSentence getSentence(int n, List<Integer> sentenceEnds, List<AnalyzedSentence> sentences) {
    for (int i = 0; i < sentenceEnds.size(); i++) {
      if (n >= i) {
        return (i == 0) ? sentences.get(0) : sentences.get(i - 1);
      }
    }
    return sentences.get(sentences.size() - 1);
  }
*/
  private boolean isCorrectSuggestion(String suggestion, WtAiToken singleWordToken) {
    return (suggestion.isEmpty() || singleWordToken == null || singleWordToken.isNonWord()
        || linguServices.isCorrectSpell(suggestion, locale));
  }

  private boolean isEndQoute(int j, List<WtAiToken> resultTokens) {
    return (j == resultTokens.size() - 1 && QUOTES.matcher(resultTokens.get(j).getToken()).matches());
  }

  private void setType(int nParaTokenStart, int nParaTokenEnd, int nResultTokenStart, int nResultTokenEnd,
      List<WtAiToken> paraTokens, List<WtAiToken> resultTokens, RuleMatch ruleMatch) {
    WtMessageHandler.printToLogFile("ParaTokenStart: " + paraTokens.get(nParaTokenStart).getToken()
        +", ParaTokenEnd: " + paraTokens.get(nParaTokenEnd).getToken()
        + ", ResultTokenStart: " + resultTokens.get(nResultTokenStart).getToken()
        +", ResultTokenEnd: " + resultTokens.get(nResultTokenEnd).getToken());
    if (nParaTokenStart == nParaTokenEnd && nResultTokenStart == nResultTokenEnd) {
      if (PUNCTUATION.matcher(paraTokens.get(nParaTokenStart).getToken()).matches()) {
        ruleMatch.setType(Type.Hint);
      } else if (shareLemma(paraTokens.get(nParaTokenStart), resultTokens.get(nResultTokenStart))
          || isHintException(paraTokens.get(nParaTokenStart), resultTokens.get(nResultTokenStart))) {
        ruleMatch.setType(Type.Hint);
      } else {
        ruleMatch.setType(Type.Other);
      } 
    } else if (nParaTokenStart + 1 == nParaTokenEnd && nResultTokenStart == nResultTokenEnd
        && ((PUNCTUATION.matcher(paraTokens.get(nParaTokenStart).getToken()).matches() 
              && paraTokens.get(nParaTokenEnd).getToken().equals(resultTokens.get(nResultTokenStart).getToken()))
            || (PUNCTUATION.matcher(paraTokens.get(nParaTokenEnd).getToken()).matches()
              && paraTokens.get(nParaTokenStart).getToken().equals(resultTokens.get(nResultTokenStart).getToken())))) {
      ruleMatch.setType(Type.Hint);
    } else if (nParaTokenStart == nParaTokenEnd && nResultTokenStart + 1 == nResultTokenEnd
        && ((PUNCTUATION.matcher(resultTokens.get(nResultTokenStart).getToken()).matches() 
              && resultTokens.get(nResultTokenEnd).getToken().equals(paraTokens.get(nParaTokenStart).getToken()))
            || (PUNCTUATION.matcher(resultTokens.get(nResultTokenEnd).getToken()).matches()
              && resultTokens.get(nResultTokenStart).getToken().equals(paraTokens.get(nParaTokenStart).getToken())))) {
      ruleMatch.setType(Type.Hint);
    } else {
      ruleMatch.setType(Type.Other);
    }
  }
  
  private boolean matchesShareLemma(int nParaTokenStart, int nParaTokenEnd, int nResultTokenStart, int nResultTokenEnd,
      List<WtAiToken> paraTokens, List<WtAiToken> resultTokens) {
    for (int i = nParaTokenStart; i <= nParaTokenEnd; i++) {
      for (int j = nResultTokenStart; j <= nResultTokenStart; j++) {
        if (shareLemma(paraTokens.get(i), resultTokens.get(j))) {
          return true;
        }
      }
    }
    return false;
  }
  
  private boolean styleHintAssumed(int nRuleTokens, int nSentTokens, 
      List<AiRuleMatch> aiMatches, List<WtAiToken> paraTokens, List<WtAiToken> resultTokens) {
    if (nRuleTokens >= nSentTokens / 2) {
      return true;
    }
    if (nSentTokens > 300 / STYLE_HINT_LIMIT && nRuleTokens >= 100 / STYLE_HINT_LIMIT ) {
      return true;
    }
    for (int i = 0; i < aiMatches.size(); i++) {
      for (int j = 0; j < aiMatches.size(); j++) {
        if (i != j) {
          if (matchesShareLemma(aiMatches.get(i).nParaTokenStart, aiMatches.get(i).nParaTokenEnd, 
              aiMatches.get(j).nResultTokenStart, aiMatches.get(j).nResultTokenEnd,
              paraTokens, resultTokens)) {
            return true;
          }
        }
      }
    }
    return false;
  }
  
  private boolean shareLemma(WtAiToken a, WtAiToken b) {
    for (AnalyzedToken t : a.getReadings()) {
      String lemma = t.getLemma();
      if (lemma != null && !lemma.isEmpty() && b.hasLemma(lemma)) {
        return true;
      }
    }
    return false;
  }
  
  private void addAllRuleMatches(List<RuleMatch> matches, List<AiRuleMatch> aiMatches) {
    for (AiRuleMatch match : aiMatches) {
      matches.add(match.ruleMatch);
    }
  }
  
  /**
   * Set language specific exceptions to set Color for specific Languages
   */
  public boolean isHintException(WtAiToken paraToken, WtAiToken resultToken) {
    return false;   
  }

  /**
   * Set language specific exceptions to handle change as a match
   */
  public boolean isMatchException(int nParaStart, int nParaEnd, 
      int nResultStart, int nResultEnd, List<WtAiToken> paraTokens, List<WtAiToken> resultTokens) {
    return false;   
  }

  @Override
  public int minToCheckParagraph() {
    return 0;
  }

  @Override
  public String getId() {
    return WtOfficeTools.AI_GRAMMAR_RULE_ID;
  }

  @Override
  public String getDescription() {
    return messages.getString("loAiRuleDescription");
  }
  
  class AiRuleMatch {
    public final RuleMatch ruleMatch;
    public final int suggestionStart;
    public final int suggestionEnd;
    public final int nParaTokenStart;
    public final int nParaTokenEnd;
    public final int nResultTokenStart;
    public final int nResultTokenEnd;
    
    AiRuleMatch(RuleMatch ruleMatch, int suggestionStart, int suggestionEnd, 
        int nParaTokenStart, int nParaTokenEnd, int nResultTokenStart, int nResultTokenEnd) {
      this.ruleMatch = ruleMatch;
      this.suggestionStart = suggestionStart;
      this.suggestionEnd = suggestionEnd;
      this.nParaTokenStart = nParaTokenStart;
      this.nParaTokenEnd = nParaTokenEnd;
      this.nResultTokenStart = nResultTokenStart;
      this.nResultTokenEnd = nResultTokenEnd;
    }
  }

}

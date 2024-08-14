package org.writingtool.aisupport;

import java.util.List;
import java.util.ResourceBundle;

import org.languagetool.AnalyzedSentence;
import org.writingtool.WtLinguisticServices;
import com.sun.star.lang.Locale;

public class WtAiDetectionRule_de extends WtAiDetectionRule {

  WtAiDetectionRule_de(String aiResultText, List<AnalyzedSentence> analyzedAiResult, WtLinguisticServices linguServices,
      Locale locale, ResourceBundle messages, boolean showStylisticHints) {
    super(aiResultText, analyzedAiResult, linguServices, locale, messages, showStylisticHints);
  }
  
  @Override
  public String getLanguage() {
    return "de";
  }

  /**
   * Set Exceptions to set Color for specific Languages
   */
  @Override
  public boolean isHintException(WtAiToken paraToken, WtAiToken resultToken) {
//    WtMessageHandler.printToLogFile("isHintException in: de" 
//        + ", paraToken: " + paraToken.getToken() + ", resultToken: " + resultToken.getToken());
    if ("dass".equals(paraToken.getToken())  || "dass".equals(resultToken.getToken())) {
      return true;
    }
    return false;   
  }

  /**
   * If tokens (from start to end) contains sToken return true 
   * else false
   */
  private boolean containToken(String sToken, int start, int end, List<WtAiToken> tokens) {
    for (int i = start; i <= end; i++) {
      if (sToken.equals(tokens.get(i).getToken())) {
        return true;
      }
    }
    return false;
  }

  /**
   * Set language specific exceptions to handle change as a match
   */
  @Override
  public boolean isMatchException(int nParaStart, int nParaEnd,
      int nResultStart, int nResultEnd, List<WtAiToken> paraTokens, List<WtAiToken> resultTokens) {
    if (nResultStart < 0 || nResultStart >= resultTokens.size() - 1) {
      return false;
    }
    if ((resultTokens.get(nResultStart).getToken().equals(",") && nResultStart < resultTokens.size() - 1 
          && (resultTokens.get(nResultStart + 1).getToken().equals("und") || resultTokens.get(nResultStart + 1).getToken().equals("oder")))
        || (resultTokens.get(nResultStart + 1).getToken().equals(",") && nResultStart < resultTokens.size() - 2 
            && (resultTokens.get(nResultStart + 2).getToken().equals("und") 
                || resultTokens.get(nResultStart + 2).getToken().equals("oder")))) {
      return true;
    }
    if (nParaStart < paraTokens.size() - 1
        && isQuote(paraTokens.get(nParaStart).getToken()) 
        && ",".equals(paraTokens.get(nParaStart + 1).getToken())
        && !containToken(paraTokens.get(nParaStart).getToken(), nResultStart, nResultEnd, resultTokens)) {
      return true;
    }
    return false;   
  }



}

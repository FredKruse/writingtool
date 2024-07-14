package org.writingtool.aisupport;

import java.util.List;
import java.util.ResourceBundle;

import org.languagetool.AnalyzedSentence;
import org.writingtool.WtLinguisticServices;
import org.writingtool.tools.WtMessageHandler;

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
   * Set language specific exceptions to handle change as a match
   */
  @Override
  public boolean isMatchException(int nPara, int nResult, List<WtAiToken> paraTokens, List<WtAiToken> resultTokens) {
    if (nResult < 0 || nResult >= resultTokens.size() - 1) {
      return false;
    }
    if ((resultTokens.get(nResult).getToken().equals(",") && nResult < resultTokens.size() - 1 
          && (resultTokens.get(nResult + 1).getToken().equals("und") || resultTokens.get(nResult + 1).getToken().equals("oder")))
        || (resultTokens.get(nResult + 1).getToken().equals(",") && nResult < resultTokens.size() - 2 
            && (resultTokens.get(nResult + 2).getToken().equals("und") || resultTokens.get(nResult + 2).getToken().equals("oder")))) {
      return true;
    }
    return false;   
  }



}

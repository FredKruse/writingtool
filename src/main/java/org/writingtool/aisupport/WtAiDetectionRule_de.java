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
    WtMessageHandler.printToLogFile("isHintException in: de" 
        + ", paraToken: " + paraToken.getToken() + ", resultToken: " + resultToken.getToken());
    if ("dass".equals(paraToken.getToken())  || "dass".equals(resultToken.getToken())) {
      return true;
    }
    return false;   
  }



}

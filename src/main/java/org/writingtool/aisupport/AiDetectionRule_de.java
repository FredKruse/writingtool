package org.writingtool.aisupport;

import java.util.List;
import java.util.ResourceBundle;

import org.languagetool.AnalyzedSentence;
import org.writingtool.LinguisticServices;
import org.writingtool.MessageHandler;

import com.sun.star.lang.Locale;

public class AiDetectionRule_de extends AiDetectionRule {

  AiDetectionRule_de(String aiResultText, List<AnalyzedSentence> analyzedAiResult, LinguisticServices linguServices,
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
  public boolean isHintException(AiToken paraToken, AiToken resultToken) {
    MessageHandler.printToLogFile("isHintException in: de" 
        + ", paraToken: " + paraToken.getToken() + ", resultToken: " + resultToken.getToken());
    if ("dass".equals(paraToken.getToken())  || "dass".equals(resultToken.getToken())) {
      return true;
    }
    return false;   
  }



}

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

import java.util.List;

import org.languagetool.AnalyzedSentence;
import org.languagetool.AnalyzedToken;
import org.languagetool.AnalyzedTokenReadings;

/**
 * Defines an easy token
 * @since 6.5
 * @author Fred Kruse
 */
public class AiToken {
  private final AnalyzedTokenReadings token;
  private final AnalyzedSentence sentence;
  private final int sentencePos;
  
  AiToken(AnalyzedTokenReadings token, int sentencePos, AnalyzedSentence sentence) {
    this.sentence = sentence;
    this.token = token;
    this.sentencePos = sentencePos;
  }
  
  int getStartPos() {
    return token.getStartPos() + sentencePos;
  }

  int getEndPos() {
    return token.getEndPos() + sentencePos;
  }

  String getToken() {
    return token.getToken();
  }
  
  AnalyzedSentence getSentence() {
    return sentence;
  }
  
  boolean isNonWord() {
    return token.isNonWord();
  }
  
  boolean hasLemma(String s) {
    return token.hasLemma(s);
  }
  
  List<AnalyzedToken> getReadings() {
    return token.getReadings();
  }
 
}

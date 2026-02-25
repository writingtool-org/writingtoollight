/* WritingTool, a LibreOffice Extension based on LanguageTool
 * Copyright (C) 2024 Fred Kruse (https://writingtool.org)
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

import java.awt.Color;
import java.net.URL;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.languagetool.JLanguageTool;
import org.languagetool.Language;
import org.languagetool.rules.RuleMatch;
import org.languagetool.tools.StringTools;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtOfficeTools.LoErrorType;

import com.sun.star.beans.PropertyState;
import com.sun.star.beans.PropertyValue;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XComponent;
import com.sun.star.linguistic2.ProofreadingResult;
import com.sun.star.linguistic2.SingleProofreadingError;
import com.sun.star.text.TextMarkupType;
import com.sun.star.uno.XComponentContext;

/**
 * Class for checking text of one LO document 
 */
public class WtSingleDocument {
  
  /**
   * Full text Check:
   * numParasToCheck: Paragraphs to be checked for full text rules
   * < 0 check full text (time intensive)
   * == 0 check only one paragraph (works like LT Version <= 3.9)
   * > 0 checks numParasToCheck before and after the processed paragraph
   * 
   * Cache:
   * sentencesCache: only used for doResetCheck == true (LO checks again only changed paragraphs by default)
   * paragraphsCache: used to store LT matches for a fast return to LO (numParasToCheck != 0)
   * singleParaCache: used for one paragraph check by default or for special paragraphs like headers, footers, footnotes, etc.
   *  
   */
  
  private static int debugMode;                   //  should be 0 except for testing; 1 = low level; 2 = advanced level
  private static boolean debugModeTm;             // time measurement should be false except for testing
  
  private WtConfiguration config;

  private String docID;                           //  docID of the document
  private XComponent xComponent;                  //  XComponent of the open document
  private final WtDocumentsHandler mDocHandler;      //  handles the different documents loaded in LO/OO
  
  private Language docLanguage;                   //  docLanguage (usually the Language of the first paragraph)

  WtSingleDocument(XComponentContext xContext, WtConfiguration config, String docID, 
      XComponent xComp, WtDocumentsHandler mDH, Language lang) {
    debugMode = WtOfficeTools.DEBUG_MODE_SD;
    debugModeTm = WtOfficeTools.DEBUG_MODE_TM;
    this.config = config;
    this.docID = docID;
    docLanguage = lang;
    xComponent = xComp;
    mDocHandler = mDH;
    if (config != null) {
      setConfigValues(config);
    }
  }
  
  /**  get the result for a check of a single document 
   * 
   * @param paraText          paragraph text
   * @param paRes             proof reading result
   * @return                  proof reading result
   */
  ProofreadingResult getCheckResults(String paraText, Locale locale, ProofreadingResult paRes, 
      PropertyValue[] propertyValues, boolean docReset, WtLanguageTool lt, LoErrorType errType) {
    return getCheckResults(paraText, locale, paRes, propertyValues, docReset, lt, -1, errType);
  }
    
  public ProofreadingResult getCheckResults(String paraText, Locale locale, ProofreadingResult paRes, 
      PropertyValue[] propertyValues, boolean docReset, WtLanguageTool lt, int nPara, LoErrorType errType) {
    try {
      int [] footnotePositions = null;  // e.g. for LO/OO < 4.3 and the 'FootnotePositions' property
      for (PropertyValue propertyValue : propertyValues) {
        if ("FootnotePositions".equals(propertyValue.Name)) {
          if (propertyValue.Value instanceof int[]) {
            footnotePositions = (int[]) propertyValue.Value;
          } else {
            WtMessageHandler.printToLogFile("SingleDocument: getCheckResults: Not of expected type int[]: " + propertyValue.Name + ": " + propertyValue.Value.getClass());
          }
        }
      }
      if (docLanguage == null) {
        docLanguage = lt.getLanguage();
      }
      long startTime = 0;
      if (debugModeTm) {
        startTime = System.currentTimeMillis();
      }
      String pText = removeFootnotes(paraText, footnotePositions);
      List<RuleMatch> paragraphMatches = lt.check(pText, JLanguageTool.ParagraphHandling.NORMAL);
      if (paragraphMatches == null || paragraphMatches.isEmpty()) {
        paRes.aErrors = new SingleProofreadingError[0];
        if (debugMode > 1) {
          WtMessageHandler.printToLogFile("WtSingleDocument: getCheckResults: Errors: 0");
        }
      } else {
        List<SingleProofreadingError> errorList = new ArrayList<>();
        for (RuleMatch myRuleMatch : paragraphMatches) {
          int toPos = myRuleMatch.getToPos();
          if (toPos > paraText.length()) {
            toPos = paraText.length();
          }
          SingleProofreadingError error = correctRuleMatchWithFootnotes(
              createOOoError(myRuleMatch, 0, footnotePositions, docLanguage, config), footnotePositions);
          errorList.add(error);
        }
        paRes.aErrors = errorList.toArray(new SingleProofreadingError[0]);
      }
      paRes.nStartOfNextSentencePosition = getNextSentencePosition(paraText, lt, paRes.nStartOfSentencePosition);
      if (paRes.nStartOfNextSentencePosition == 0) {
        paRes.nStartOfNextSentencePosition = paraText.length();
      }
      paRes.nBehindEndOfSentencePosition = paRes.nStartOfNextSentencePosition;
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
        if (runTime > WtOfficeTools.TIME_TOLERANCE) {
          WtMessageHandler.printToLogFile("Single document: Time to run single check: " + runTime);
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return paRes;
  }
  
  /**
   * set values set by configuration dialog
   */
  void setConfigValues(WtConfiguration config) {
    this.config = config;
  }

  /**
   * get language of the document
   */
  public Language getLanguage() {
    return docLanguage;
  }
  
  /**
   * set language of the document
   */
  void setLanguage(Language language) {
    docLanguage = language;
  }
  
  /** 
   * Set XComponentContext and XComponent of the document
   */
  void setXComponent(XComponentContext xContext, XComponent xComponent) {
    this.xComponent = xComponent;
  }
  
  /**
   *  Get xComponent of the document
   */
  public XComponent getXComponent() {
    return xComponent;
  }
  
  /**
   *  Get MultiDocumentsHandler
   */
  public WtDocumentsHandler getMultiDocumentsHandler() {
    return mDocHandler;
  }
  
  /**
   *  Get ID of the document
   */
  public String getDocID() {
    return docID;
  }
  
  /**
   *  Set ID of the document
   */
  void setDocID(String docId) {
    docID = docId;
  }
  
  /**
   * reset the Document
   */
  void resetDocument() {
    mDocHandler.resetDocument();
  }

  /**
   * get all synonyms as array
   */
  public String[] getSynonymArray(SingleProofreadingError error, String para, Locale locale, WtLanguageTool lt, boolean setLimit) {
    Map<String, List<String>> synonymMap = getSynonymMap(error, para, locale, lt);
    if (synonymMap.isEmpty()) {
      return new String[0];
    }
    List<String> suggestions = new ArrayList<>();
    int n = 0;
    for (String lemma : synonymMap.keySet()) {
      for (String suggestion : synonymMap.get(lemma)) {
        suggestions.add(suggestion);
        n++;
        if (setLimit && n >= WtOfficeTools.MAX_SUGGESTIONS) {
          break;
        }
      }
      if (setLimit && n >= WtOfficeTools.MAX_SUGGESTIONS) {
        break;
      }
    }
    return suggestions.toArray(new String[suggestions.size()]);
  }
  
  /**
   * get all synonyms as map
   */
  public Map<String, List<String>> getSynonymMap(SingleProofreadingError error, String para, Locale locale, WtLanguageTool lt) {
    Map<String, List<String>> suggestionMap = new HashMap<>();
    try {
      String word = para.substring(error.nErrorStart, error.nErrorStart + error.nErrorLength);
      boolean startUpperCase = Character.isUpperCase(word.charAt(0));
      if (debugMode > 0) {
        WtMessageHandler.printToLogFile("SingleDocument: getSynonymMap: Find Synonyms for word:" + word);
      }
//      List<String> lemmas = lt.getLemmasOfWord(word);
      List<String> lemmas = lt.isRemote() ? lt.getLemmasOfWord(word) : lt.getLemmasOfParagraph(para, error.nErrorStart);
      for (String lemma : lemmas) {
        if (debugMode > 1) {
          WtMessageHandler.printToLogFile("SingleDocument: getSynonymMap: Find Synonyms for lemma:" + lemma);
        }
        List<String> suggestions = new ArrayList<>();
        List<String> synonyms = mDocHandler.getLinguisticServices().getSynonyms(lemma, locale);
        for (String synonym : synonyms) {
          synonym = synonym.replaceAll("\\(.*\\)", "").trim();
          if (debugMode > 1) {
            WtMessageHandler.printToLogFile("SingleDocument: getSynonymMap: Synonym:" + synonym);
          }
          if (!synonym.isEmpty() && !suggestions.contains(synonym)
              && ( (startUpperCase && Character.isUpperCase(synonym.charAt(0))) 
                  || (!startUpperCase && Character.isLowerCase(synonym.charAt(0))))) {
            suggestions.add(synonym);
          }
        }
        if (!suggestions.isEmpty()) {
          suggestionMap.put(lemma, suggestions);
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
    return suggestionMap;
  }

  /**
   * Fix numbers that are (probably) foot notes.
   * See https://bugs.freedesktop.org/show_bug.cgi?id=69416
   * public for test reasons
   */
  static String cleanFootnotes(String paraText) {
    return paraText.replaceAll("([^\\d][.!?])\\d ", "$1¹ ");
  }
  
  /**
   * Remove footnotes, deleted characters and hidden characters from paraText
   * run cleanFootnotes if information about footnotes are not supported
   */
  public static String removeFootnotes(String paraText, int[] footnotes) throws Throwable {
    if (paraText == null) {
      return null;
    }
    if (footnotes == null) {
      return cleanFootnotes(paraText);
    }
    for (int i = footnotes.length - 1; i >= 0; i--) {
      if (footnotes[i] < paraText.length()) {
        paraText = paraText.substring(0, footnotes[i]) + paraText.substring(footnotes[i] + 1);
      }
    }
    return paraText;
  }

  private static PropertyValue createPropertyValue(String name, Object value) {
    PropertyValue property = new PropertyValue();
    property.Name = name;
    property.Value = value;
    property.Handle = -1;
    property.State = PropertyState.DIRECT_VALUE;
    return property;
  }

  /**
   * get beginning of next sentence using LanguageTool tokenization
   */
  public static int getNextSentencePosition(String paraText, WtLanguageTool lt, int startPos) {
    if (paraText == null || paraText.isEmpty()) {
      return 0;
    }
    if (lt == null) {
      return paraText.length();
    } else {
      List<String> tokenizedSentences = lt.sentenceTokenize(cleanFootnotes(paraText));
      int position = 0;
      for (String sentence : tokenizedSentences) {
        position += sentence.length();
        if (position > startPos) {
          return position;
        }
      }
    }
    return paraText.length();
  }

  /**
   * get beginning of next sentence using LanguageTool tokenization
   */
  public static List<Integer> getNextSentencePositions (String paraText, WtLanguageTool lt) {
    List<Integer> nextSentencePositions = new ArrayList<Integer>();
    if (paraText == null || paraText.isEmpty()) {
      nextSentencePositions.add(0);
      return nextSentencePositions;
    }
//    if (lt == null || lt.isRemote()) {
    if (lt == null) {
      nextSentencePositions.add(paraText.length());
    } else {
      List<String> tokenizedSentences = lt.sentenceTokenize(cleanFootnotes(paraText));
      int position = 0;
      for (String sentence : tokenizedSentences) {
        position += sentence.length();
        nextSentencePositions.add(position);
      }
      if (nextSentencePositions.get(nextSentencePositions.size() - 1) != paraText.length()) {
        nextSentencePositions.set(nextSentencePositions.size() - 1, paraText.length());
      }
    }
    return nextSentencePositions;
  }
  
  /**
   * Correct WtProofreadingError by footnote positions
   * footnotes before is the sum of all footnotes before the checked paragraph
   */
  public static SingleProofreadingError correctRuleMatchWithFootnotes(SingleProofreadingError pError, int[] footnotes) throws Throwable {
    if (footnotes == null) {
      return pError;
    }
    for (int i : footnotes) {
      if (i <= pError.nErrorStart) {
        pError.nErrorStart++;
      } else if (i < pError.nErrorStart + pError.nErrorLength) {
        pError.nErrorLength++;
      }
    }
    return pError;
  }

  /**
   * Creates a SingleGrammarError object for use in LO/OO.
   */
  public static SingleProofreadingError createOOoError(RuleMatch ruleMatch, int startIndex, int[] footnotes,
      Language docLanguage, WtConfiguration config) throws Throwable {
    SingleProofreadingError aError = new SingleProofreadingError();
    if (ruleMatch.getRule().isDictionaryBasedSpellingRule()) {
      aError.nErrorType = TextMarkupType.SPELLCHECK;
    } else {
      aError.nErrorType = TextMarkupType.PROOFREADING;
    }
    
    // the API currently has no support for formatting text in comments
    String msg = ruleMatch.getMessage();
    if (docLanguage != null) {
      msg = docLanguage.toAdvancedTypography(msg);
    }
    msg = msg.replaceAll("<suggestion>", docLanguage == null ? "\"" : docLanguage.getOpeningDoubleQuote())
        .replaceAll("</suggestion>", docLanguage == null ? "\"" : docLanguage.getClosingDoubleQuote())
        .replaceAll("([\r]*\n)", " "); 
    aError.aFullComment = msg;
    // not all rules have short comments
    if (!config.useLongMessages() && !StringTools.isEmpty(ruleMatch.getShortMessage())) {
      aError.aShortComment = ruleMatch.getShortMessage();
    } else {
      aError.aShortComment = aError.aFullComment;
    }
    aError.aShortComment = org.writingtool.tools.WtGeneralTools.shortenComment(aError.aShortComment);
    //  Filter: provide user to delete footnotes by suggestion
    boolean noSuggestions = false;
    if (footnotes != null && footnotes.length > 0 && !ruleMatch.getSuggestedReplacements().isEmpty()) {
      int cor = 0;
      for (int n : footnotes) {
        if (n + cor <= ruleMatch.getFromPos() + startIndex) {
          cor++;
        } else if (n + cor > ruleMatch.getFromPos() + startIndex && n + cor <= ruleMatch.getToPos() + startIndex) {
          noSuggestions = true;
          break;
        }
      }
    }
    int numSuggestions;
    String[] allSuggestions;
    if (noSuggestions) {
      numSuggestions = 0;
      allSuggestions = new String[0];
    } else {
      numSuggestions = ruleMatch.getSuggestedReplacements().size();
      allSuggestions = ruleMatch.getSuggestedReplacements().toArray(new String[numSuggestions]);
    }
    //  Filter: remove suggestions for override dot at the end of sentences
    //  needed because of error in dialog
    /*  since LT 5.2: Filter is commented out because of default use of LT dialog
    if (lastChar == '.' && (ruleMatch.getToPos() + startIndex) == sentencesLength) {
      int i = 0;
      while (i < numSuggestions && i < OfficeTools.MAX_SUGGESTIONS
          && allSuggestions[i].length() > 0 && allSuggestions[i].charAt(allSuggestions[i].length()-1) == '.') {
        i++;
      }
      if (i < numSuggestions && i < OfficeTools.MAX_SUGGESTIONS) {
      numSuggestions = 0;
      allSuggestions = new String[0];
      }
    }
    */
    //  End of Filter
    if (numSuggestions > WtOfficeTools.MAX_SUGGESTIONS) {
      aError.aSuggestions = Arrays.copyOfRange(allSuggestions, 0, WtOfficeTools.MAX_SUGGESTIONS);
    } else {
      aError.aSuggestions = allSuggestions;
    }
    aError.nErrorStart = ruleMatch.getFromPos() + startIndex;
    aError.nErrorLength = ruleMatch.getToPos() - ruleMatch.getFromPos();
    aError.aRuleIdentifier = ruleMatch.getRule().getId();
    // LibreOffice since version 3.5 supports an URL that provides more information about the error,
    // LibreOffice since version 6.2 supports the change of underline color (key: "LineColor", value: int (RGB))
    // LibreOffice since version 6.2 supports the change of underline style (key: "LineType", value: short (DASHED = 5))
    // older version will simply ignore the properties
    String category = ruleMatch.getRule().getCategory().getName();
    String ruleId = ruleMatch.getRule().getId();
    Color underlineColor = aError.nErrorType != TextMarkupType.SPELLCHECK 
                            ? config.getUnderlineColor(category, ruleId) : Color.red;
    short underlineType = aError.nErrorType != TextMarkupType.SPELLCHECK 
                            ? config.getUnderlineType(category, ruleId) : WtConfiguration.UNDERLINE_WAVE;
    URL url = ruleMatch.getUrl();
    if (url == null) {                      // match URL overrides rule URL 
      url = ruleMatch.getRule().getUrl();
    }
    int nDim = 0;
    if (url != null) {
      nDim++;
    }
    if (underlineColor != Color.blue || aError.nErrorType == TextMarkupType.SPELLCHECK) {
      nDim++;
    }
    if (underlineType != WtConfiguration.UNDERLINE_WAVE || (config.markSingleCharBold() && aError.nErrorLength == 1)) {
      nDim++;
    }
    if (nDim > 0) {
      //  HINT: Because of result cache handling:
      //  handle should always be -1
      //  property state should always be PropertyState.DIRECT_VALUE
      //  otherwise result cache handling has to be adapted
      PropertyValue[] propertyValues = new PropertyValue[nDim];
      int n = 0;
      if (url != null) {
        propertyValues[n] = createPropertyValue("FullCommentURL", url.toString());
        n++;
      }
      if (aError.nErrorType == TextMarkupType.SPELLCHECK) {
        int ucolor = Color.red.getRGB() & 0xFFFFFF;
        propertyValues[n] = createPropertyValue("LineColor", ucolor);
        n++;
/*        
      } else if (ruleMatch.getRule() instanceof WtAiDetectionRule) {
        int ucolor;
        if (ruleMatch.getType() == Type.Hint) {
          ucolor = WtAiDetectionRule.RULE_HINT_COLOR.getRGB() & 0xFFFFFF;
        } else {
          ucolor = WtAiDetectionRule.RULE_OTHER_COLOR.getRGB() & 0xFFFFFF;
        }
        propertyValues[n] = new PropertyValue("LineColor", -1, ucolor, PropertyState.DIRECT_VALUE);
        n++;
*/
      } else {
        if (underlineColor != Color.blue) {
          int ucolor = underlineColor.getRGB() & 0xFFFFFF;
          propertyValues[n] = createPropertyValue("LineColor", ucolor);
          n++;
        }
      }
      if (underlineType != WtConfiguration.UNDERLINE_WAVE) {
        propertyValues[n] = createPropertyValue("LineType", underlineType);
      } else if (config.markSingleCharBold() && aError.nErrorLength == 1) {
        propertyValues[n] = createPropertyValue("LineType", WtConfiguration.UNDERLINE_BOLDWAVE);
      }
      aError.aProperties = propertyValues;
    } else {
        aError.aProperties = new PropertyValue[0];
    }
    return aError;
  }
 
}

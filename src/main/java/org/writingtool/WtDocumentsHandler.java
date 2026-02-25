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

import java.awt.*;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.io.File;
import java.io.IOException;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.ResourceBundle;
import java.util.Set;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.UIManager;

import org.languagetool.Language;
import org.languagetool.Languages;
import org.languagetool.UserConfig;
import org.languagetool.rules.Rule;
import org.languagetool.tools.Tools;
import org.writingtool.config.WtConfiguration;
import org.writingtool.dialogs.WtAboutDialog;
import org.writingtool.dialogs.WtConfigurationDialog;
import org.writingtool.config.WtConfigThread;
import org.writingtool.tools.WtGeneralTools;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtOfficeTools.DocumentType;
import org.writingtool.tools.WtOfficeTools.LoErrorType;
import org.writingtool.tools.WtVersionInfo;

import com.sun.star.beans.PropertyValue;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XEventListener;
import com.sun.star.linguistic2.LinguServiceEvent;
import com.sun.star.linguistic2.LinguServiceEventFlags;
import com.sun.star.linguistic2.ProofreadingResult;
import com.sun.star.linguistic2.XLinguServiceEventListener;
import com.sun.star.linguistic2.XProofreader;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


/**
 * Class to handle multiple LO documents for checking
 * @since 1.0
 * @author Fred Kruse, Marcin Miłkowski
 */
public class WtDocumentsHandler {

  // LibreOffice (since 4.2.0) special tag for locale with variant 
  // e.g. language ="qlt" country="ES" variant="ca-ES-valencia":
  private static final String LIBREOFFICE_SPECIAL_LANGUAGE_TAG = "qlt";
  
  private static final ResourceBundle messages = WtOfficeTools.getMessageBundle();

  private final List<XLinguServiceEventListener> xEventListeners;
  private boolean docReset = false;

  private static boolean debugMode = false;   //  should be false except for testing
  private static boolean debugModeTm = false; //  should be false except for testing

  public final boolean isOpenOffice;
  
  private WtLanguageTool lt = null;
  private Language docLanguage = null;
  private Language fixedLanguage = null;
  private Language langForShortName;
  private Locale locale;                            //  locale for grammar check
  private final XEventListener xEventListener;
  private final XProofreader xProofreader;
  private static File configDir;
  private static String configFile;
  private static WtConfiguration config = null;
  private WtLinguisticServices linguServices = null;
  private static Map<String, Set<String>> disabledRulesUI;  //  Rules disabled by context menu or spell dialog
  private final List<Rule> extraRemoteRules;                //  store of rules supported by remote server but not locally
  private WtConfigurationDialog cfgDialog = null;           //  configuration dialog (show only one configuration panel)
  private static WtAboutDialog aboutDialog = null;          //  about dialog (show only one about panel)
  
  private static XComponentContext xContext;          //  The context of the document
  private final List<WtSingleDocument> documents;     //  The List of LO documents to be checked
  private final List<String> disposedIds;             //  The List of IDs of disposed documents
  private boolean recheck = true;                     //  if true: recheck the whole document at next iteration
  private int docNum;                                 //  number of the current document
  

  private boolean noBackgroundCheck = false;          //  is LT switched off by config
  private boolean noLtSpeller = false;                //  true if LT spell checker can't be used

  private boolean useOrginalCheckDialog = false;      // use original spell and grammar dialog (LT check dialog does not work for OO)
  private boolean checkImpressDocument = false;       //  the document to check is Impress
  private boolean isNotTextDocument = false;
  private boolean testMode = false;
  private static int javaLookAndFeelSet = -1;

  
  WtDocumentsHandler(XComponentContext xContext, XProofreader xProofreader, XEventListener xEventListener) {
    WtDocumentsHandler.xContext = xContext;
    this.xEventListener = xEventListener;
    this.xProofreader = xProofreader;
    xEventListeners = new ArrayList<>();
    WtVersionInfo.init(xContext);
    WtOfficeTools.renameOldLtFiles();     // This has to be deleted in later versions 
    if (WtVersionInfo.ooName == null || WtVersionInfo.ooName.equals("OpenOffice")) {
      isOpenOffice = true;
      useOrginalCheckDialog = true;
      configFile = WtOfficeTools.OOO_CONFIG_FILE;
    } else {
      isOpenOffice = false;
      configFile = WtOfficeTools.CONFIG_FILE;
    }
    configDir = WtOfficeTools.getWtConfigDir(xContext);
    WtMessageHandler.init(xContext, false);
    documents = new ArrayList<>();
    disposedIds = new ArrayList<>();
    disabledRulesUI = new HashMap<>();
    extraRemoteRules = new ArrayList<>();
    if (WtVersionInfo.osArch == null || WtVersionInfo.osArch.equals("x86")
        || !WtSpellChecker.runLTSpellChecker(xContext)) {
      noLtSpeller = true;
    }
  }
  
  /**
   * Runs the grammar checker on paragraph text.
   *
   * @param docID document ID
   * @param paraText paragraph text
   * @param locale Locale the text Locale
   * @param startOfSentencePos start of sentence position
   * @param nSuggestedBehindEndOfSentencePosition end of sentence position
   * @return ProofreadingResult containing the results of the check.
   */
  public final ProofreadingResult doProofreading(String docID,
      String paraText, Locale locale, int startOfSentencePos,
      int nSuggestedBehindEndOfSentencePosition,
      PropertyValue[] propertyValues) {
    ProofreadingResult paRes = new ProofreadingResult();
    paRes.nStartOfSentencePosition = startOfSentencePos;
    paRes.nStartOfNextSentencePosition = nSuggestedBehindEndOfSentencePosition;
    paRes.nBehindEndOfSentencePosition = paRes.nStartOfNextSentencePosition;
    paRes.xProofreader = xProofreader;
    paRes.aLocale = locale;
    paRes.aDocumentIdentifier = docID;
    paRes.aText = paraText;
    paRes.aProperties = propertyValues;
    try {
      paRes = getCheckResults(paraText, locale, paRes, propertyValues, docReset);
      docReset = false;
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return paRes;
  }

  /**
   * distribute the check request to the concerned document
   */
  ProofreadingResult getCheckResults(String paraText, Locale locale, ProofreadingResult paRes, 
      PropertyValue[] propertyValues, boolean docReset) throws Throwable {
    if (!hasLocale(locale)) {
      docLanguage = null;
      WtMessageHandler.printToLogFile("MultiDocumentsHandler: getCheckResults: Sorry, don't have locale: " + WtOfficeTools.localeToString(locale));
      return paRes;
    }
//    LinguisticServices.setLtAsSpellService(xContext, true);
    if (!isBackgroundCheckOff()) {
      boolean isSameLanguage = true;
      if (fixedLanguage == null || langForShortName == null) {
        langForShortName = getLanguage(locale);
        isSameLanguage = langForShortName.equals(docLanguage) && lt != null;
      }
      if (!isSameLanguage || recheck || checkImpressDocument) {
        boolean initDocs = (lt == null || recheck || checkImpressDocument);
        if (debugMode && initDocs) {
          WtMessageHandler.printToLogFile("initDocs: lt " + (lt == null ? "=" : "!") + "= null, recheck: " + recheck 
              + ", Impress: " + checkImpressDocument);
        }
        checkImpressDocument = false;
        if (!isSameLanguage) {
          docLanguage = langForShortName;
          this.locale = locale;
          extraRemoteRules.clear();
        }
        lt = initLanguageTool();
        if (javaLookAndFeelSet < 0) {
          setJavaLookAndFeel();
        }
        if (initDocs) {
          initDocuments(true);
        }
      }
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("MultiDocumentsHandler: getCheckResults: Start getNumDoc!");
    }
    docNum = getNumDoc(paRes.aDocumentIdentifier, propertyValues);
    if (isBackgroundCheckOff()) {
      return paRes;
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("MultiDocumentsHandler: getCheckResults: Start testHeapSpace!");
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("MultiDocumentsHandler: getCheckResults: Start getCheckResults at single document: " + paraText);
    }
//    handleLtDictionary(paraText);
    paRes = documents.get(docNum).getCheckResults(paraText, locale, paRes, propertyValues, docReset, lt, LoErrorType.GRAMMAR);
    if (!isBackgroundCheckOff() && lt.doReset()) {
      // langTool.doReset() == true: if server connection is broken ==> switch to internal check
      WtMessageHandler.showMessage(messages.getString("loRemoteSwitchToLocal"));
      config.setRemoteCheck(false);
      try {
        config.saveConfiguration(docLanguage);
      } catch (IOException e) {
        WtMessageHandler.showError(e);
      }
      resetDocument();
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("MultiDocumentsHandler: getCheckResults: return to LO/OO!");
    }
    return paRes;
  }

  /**
   *  Get the current used document
   */
  public WtSingleDocument getCurrentDocument() {
    try {
      XComponent xComponent = WtOfficeTools.getCurrentComponent(xContext);
      isNotTextDocument = false;
      if (xComponent != null) {
        for (WtSingleDocument document : documents) {
          if (xComponent.equals(document.getXComponent())) {
            return document;
          }
        }
        XTextDocument curDoc = UnoRuntime.queryInterface(XTextDocument.class, xComponent);
        if (curDoc == null) {
          WtMessageHandler.printToLogFile("MultiDocumentsHandler: getCurrentDocument: Is document, but not a text document!");
          isNotTextDocument = true;
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return null;
  }
  
  /**
   * return true, if a document was found but is not a text document
   */
  public boolean isNotTextDocument() {
    return isNotTextDocument;
  }
  
  /**
   * return true, if the LT spell checker is not be used
   */
  public boolean noLtSpeller() {
    return this.noLtSpeller;
  }
  
  /**
   * return true, if the document to check is an Impress document
   */
  public boolean isCheckImpressDocument() {
    return checkImpressDocument;
  }
  
  /**
   * set the checkImpressDocument flag
   */
  public void setCheckImpressDocument(boolean checkImpressDocument) {
    this.checkImpressDocument = checkImpressDocument;
  }
  
  /**
   *  Set all documents to be checked again
   */
  void setRecheck() {
    recheck = true;
  }
  
  /**
   *  Set XComponentContext
   */
  void setComponentContext(XComponentContext xContxt) {
    if (xContext != null && !xContext.equals(xContxt)) {
      setRecheck();
    }
    xContext = xContxt;
  }
  
  /**
   *  Get XComponentContext
   */
  public static XComponentContext getComponentContext() {
    return xContext;
  }
  
  /**
   *  Set pointer to configuration dialog
   */
  public void setConfigurationDialog(WtConfigurationDialog dialog) {
    cfgDialog = dialog;
  }
  
  /**
   *  Set Configuration file name
   */
  public void setConfigFileName(String name) {
    configFile = name;
  }
  
  /**
   *  close configuration dialog
   * @throws Throwable 
   */
  private void closeDialogs() throws Throwable {
    if (cfgDialog != null) {
      cfgDialog.close();
      cfgDialog = null;
    }
  }
  
  /**
   *  Set a document as closed
   */
  private void setContextOfClosedDoc(XComponent xComponent) {
    boolean found = false;
    try {
      for (WtSingleDocument document : documents) {
        if (xComponent.equals(document.getXComponent())) {
          found = true;
          disposedIds.add(document.getDocID());
        }
      }
      if (!found) {
        WtMessageHandler.printToLogFile("MultiDocumentsHandler: setContextOfClosedDoc: Error: Disposed Document not found - Cache not deleted");
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }
  
  /**
   *  Add a rule to disabled rules by context menu or spell dialog
   */
  void addDisabledRule(String langCode, String ruleId) {
    if (disabledRulesUI.containsKey(langCode)) {
      disabledRulesUI.get(langCode).add(ruleId);
    } else {
      Set<String >rulesIds = new HashSet<>();
      rulesIds.add(ruleId);
      disabledRulesUI.put(langCode, rulesIds);
    }
  }
  
  /**
   *  Remove a rule from disabled rules by spell dialog
   */
  public void removeDisabledRule(String langCode, String ruleId) {
    if (disabledRulesUI.containsKey(langCode)) {
      Set<String >rulesIds = disabledRulesUI.get(langCode);
      rulesIds.remove(ruleId);
      if (rulesIds.isEmpty()) {
        disabledRulesUI.remove(langCode);
      } else {
        disabledRulesUI.put(langCode, rulesIds);
      }
    }
  }
  
  /**
   *  remove all disabled rules by context menu or spell dialog
   */
  void resetDisabledRules() {
    disabledRulesUI = new HashMap<>();
  }
  
  /**
   *  get disabled rules for a language code by context menu or spell dialog
   */
  public static Set<String> getDisabledRules(String langCode) {
    if (langCode == null || !disabledRulesUI.containsKey(langCode)) {
      return new HashSet<String>();
    }
    return disabledRulesUI.get(langCode);
  }
  
  /**
   *  get all disabled rules
   */
  Map<String, Set<String>> getAllDisabledRules() {
    return disabledRulesUI;
  }
  
  /**
   *  get all disabled rules
   */
  void setAllDisabledRules(Map<String, Set<String>> disabledRulesUI) {
    WtDocumentsHandler.disabledRulesUI = disabledRulesUI;
  }
  
  /**
   *  get all disabled rules by context menu or spell dialog
   */
  public Map<String, String> getDisabledRulesMap(String langCode) {
    if (langCode == null) {
      langCode = WtOfficeTools.localeToString(locale);
    }
    Map<String, String> disabledRulesMap = new HashMap<>();
    if (langCode != null && lt != null && config != null) {
      List<Rule> allRules = lt.getAllRules();
      List<String> disabledRules = new ArrayList<String>(getDisabledRules(langCode));
      for (int i = disabledRules.size() - 1; i >= 0; i--) {
        String disabledRule = disabledRules.get(i);
        String ruleDesc = null;
        for (Rule rule : allRules) {
          if (disabledRule.equals(rule.getId())) {
            if (!rule.isDefaultOff() || rule.isOfficeDefaultOn()) {
              ruleDesc = rule.getDescription();
            } else {
              removeDisabledRule(langCode, disabledRule);
            }
            break;
          }
        }
        if (ruleDesc != null) {
          disabledRulesMap.put(disabledRule, ruleDesc);
        }
      }
      disabledRules = new ArrayList<String>(config.getDisabledRuleIds());
      for (int i = disabledRules.size() - 1; i >= 0; i--) {
        String disabledRule = disabledRules.get(i);
        String ruleDesc = null;
        for (Rule rule : allRules) {
          if (disabledRule.equals(rule.getId())) {
            if (!rule.isDefaultOff() || rule.isOfficeDefaultOn()) {
              ruleDesc = rule.getDescription();
            } else {
              config.removeDisabledRuleId(disabledRule);
            }
            break;
          }
        }
        if (ruleDesc != null) {
          disabledRulesMap.put(disabledRule, ruleDesc);
        }
      }
    }
    return disabledRulesMap;
  }
  
  /**
   *  set disabled rules by context menu or spell dialog
   */
  public void setDisabledRules(String langCode, Set<String> ruleIds) {
    disabledRulesUI.put(langCode, new HashSet<>(ruleIds));
  }
  
  /**
   *  get LanguageTool
   */
  public WtLanguageTool getLanguageTool() {
    if (lt == null) {
      if (docLanguage == null) {
        docLanguage = getCurrentLanguage();
      }
      lt = initLanguageTool();
    }
    return lt;
  }
  
  /**
   *  get Configuration
   */
  public WtConfiguration getConfiguration() {
    try {
      if (config == null || recheck) {
        if (docLanguage == null) {
          docLanguage = getCurrentLanguage();
        }
        initLanguageTool();
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return config;
  }
  
  /**
   *  get Configuration for language
   *  @throws IOException 
   */
  public static WtConfiguration getConfiguration(Language lang) throws IOException {
    return new WtConfiguration(configDir, configFile, lang, true);
  }
  
  private void disableLTSpellChecker(XComponentContext xContext, Language lang) {
    try {
      config.setUseLtSpellChecker(false);
      config.saveConfiguration(lang);
    } catch (IOException e) {
      WtMessageHandler.printToLogFile("Can't read configuration: LT spell checker not used!");
    }
  }

  /**
   *  get LinguisticServices
   */
  public WtLinguisticServices getLinguisticServices() {
    if (linguServices == null) {
      linguServices = new WtLinguisticServices(xContext);
      WtMessageHandler.printToLogFile("MultiDocumentsHandler: getLinguisticServices: linguServices set: is " 
            + (linguServices == null ? "" : "NOT ") + "null");
      if (noLtSpeller) {
        Tools.setLinguisticServices(linguServices);
        WtMessageHandler.printToLogFile("MultiDocumentsHandler: getLinguisticServices: linguServices set to tools");
      }
    }
    return linguServices;
  }
  
  /**
   * Allow xContext == null for test cases
   */
  void setTestMode(boolean mode) {
    testMode = mode;
    WtMessageHandler.setTestMode(mode);
    if (mode) {
      configFile = "dummy_xxxx.cfg";
      File dummy = new File(configDir, configFile);
      if (dummy.exists()) {
        dummy.delete();
      }
    }
  }

  /**
   * proofs if test cases
   */
  public boolean isTestMode() {
    return testMode;
  }

  /**
   * Checks the language under the cursor. Used for opening the configuration dialog.
   * @return the language under the visible cursor
   * if null or not supported returns the most used language of the document
   * if there is no supported language use en-US as default
   */
  public Language getCurrentLanguage() {
    Locale locale = WtOfficeTools.getCursorLocale(xContext);
    if (locale == null || locale.Language.equals(WtOfficeTools.IGNORE_LANGUAGE) || !hasLocale(locale)) {
      WtSingleDocument document = getCurrentDocument();
      if (document != null) {
        Language lang = document.getLanguage();
        if (lang != null) {
          return lang;
        }
      }
      locale = new Locale("en","US","");
    }
    return getLanguage(locale);
  }
  
  /**
   * @return true if LT supports the language of a given locale
   * @param locale The Locale to check
   */
  public final static boolean hasLocale(Locale locale) {
    try {
      for (Language element : Languages.get()) {
        if (locale.Language.equalsIgnoreCase(LIBREOFFICE_SPECIAL_LANGUAGE_TAG)
            && element.getShortCodeWithCountryAndVariant().equals(locale.Variant)) {
          return true;
        }
        if (element.getShortCode().equals(locale.Language)) {
          return true;
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return false;
  }
  
  /**
   *  Set configuration Values for all documents
   */
  private void setConfigValues(WtConfiguration config, WtLanguageTool lt) throws Throwable {
    WtDocumentsHandler.config = config;
    this.lt = lt;
    for (WtSingleDocument document : documents) {
      document.setConfigValues(config);
    }
  }

  /**
   * Get language from locale
   */
  public static Language getLanguage(Locale locale) {
    try {
      if (locale.Language.equals(WtOfficeTools.IGNORE_LANGUAGE)) {
        return null;
      }
      if (locale.Language.equalsIgnoreCase(LIBREOFFICE_SPECIAL_LANGUAGE_TAG)) {
        return Languages.getLanguageForShortCode(locale.Variant);
      } else {
        return Languages.getLanguageForShortCode(locale.Language + "-" + locale.Country);
      }
    } catch (Throwable e) {
      try {
        return Languages.getLanguageForShortCode(locale.Language);
      } catch (Throwable t) {
        return null;
      }
    }
  }

  /**
   * Get or Create a Number from docID
   * Return -1 if failed
   */
  private int getNumDoc(String docID, PropertyValue[] propertyValues) throws Throwable {
    for (int i = 0; i < documents.size(); i++) {
      if (documents.get(i).getDocID().equals(docID)) {  //  document exist
        if (!testMode && documents.get(i).getXComponent() == null) {
          XComponent xComponent = WtOfficeTools.getCurrentComponent(xContext);
          if (xComponent == null) {
            WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Error: Document (ID: " + docID + ") has no XComponent -> Internal space will not be deleted when document disposes");
          } else {
            try {
              xComponent.addEventListener(xEventListener);
            } catch (Throwable t) {
              WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Error: Document (ID: " + docID + ") has no XComponent -> Internal space will not be deleted when document disposes");
              xComponent = null;
            }
            if (xComponent != null) {
              documents.get(i).setXComponent(xContext, xComponent);
              WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Fixed: XComponent set for Document (ID: " + docID + ")");
            }
          }
        }
        for (String id : disposedIds) {
          int n = removeDoc(id);
          if (n >= 0 && n < i) {
            i--;
          }
        }
        disposedIds.clear();
        return i;
      }
    }
    //  Add new document
    XComponent xComponent = null;
    if (!testMode) {              //  xComponent == null for test cases 
      xComponent = WtOfficeTools.getCurrentComponent(xContext);
      if (xComponent == null) {
        WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Error: Document (ID: " + docID + ") has no XComponent -> Internal space will not be deleted when document disposes");
      } else {
        for (int i = 0; i < documents.size(); i++) {
          //  work around to compensate a bug at LO
          if (xComponent.equals(documents.get(i).getXComponent())) {
            WtMessageHandler.printToLogFile("Different Doc IDs, but same xComponents!");
            String oldDocId = documents.get(i).getDocID();
            documents.get(i).setDocID(docID);
            documents.get(i).setLanguage(docLanguage);
            WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Document ID corrected: old: " + oldDocId + ", new: " + docID);
            return i;
          }
        }
        try {
          xComponent.addEventListener(xEventListener);
        } catch (Throwable e) {
          WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Error: Document (ID: " + docID + ") has no XComponent -> Internal space will not be deleted when document disposes");
          xComponent = null;
        }
      }
    }
    WtSingleDocument newDocument = new WtSingleDocument(xContext, config, docID, xComponent, this, docLanguage);
    documents.add(newDocument);
    for (String id : disposedIds) {
      removeDoc(id);
    }
    disposedIds.clear();
    WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Document " + (documents.size() - 1) + " created; docID = " + docID);
    return documents.size() - 1;
  }

  /**
   * Delete a document number and all internal space
   */
  private int removeDoc(String docID) {
    try {
      for (int i = documents.size() - 1; i >= 0; i--) {
        if (!docID.equals(documents.get(i).getDocID())) {
          WtMessageHandler.printToLogFile("Disposed document " + documents.get(i).getDocID() + " removed");
          documents.remove(i);
          return (i);
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return (-1);
  }
  
  /**
   * Initialize LanguageTool
   */
  public WtLanguageTool initLanguageTool() {
    return initLanguageTool(null);
  }

  public WtLanguageTool initLanguageTool(Language currentLanguage) {
    WtLanguageTool lt = null;
    try {
      config = getConfiguration(currentLanguage == null ? docLanguage : currentLanguage);
      if (this.lt == null) {
        WtOfficeTools.setLogLevel(config.getlogLevel());
        debugMode = WtOfficeTools.DEBUG_MODE_MD;
        debugModeTm = WtOfficeTools.DEBUG_MODE_TM;
        if (!noLtSpeller && !WtSpellChecker.isEnoughHeap()) {
          noLtSpeller = true;
          disableLTSpellChecker(xContext, docLanguage);
          WtMessageHandler.showMessage(messages.getString("guiSpellCheckerWarning"));
        }
      }
      long startTime = 0;
      if (debugModeTm) {
        startTime = System.currentTimeMillis();
      }
      noBackgroundCheck = config.noBackgroundCheck();
      if (linguServices == null) {
        linguServices = getLinguisticServices();
      }
      linguServices.setNoSynonymsAsSuggestions(config.noSynonymsAsSuggestions() || testMode);
      if (currentLanguage == null) {
        fixedLanguage = config.getDefaultLanguage();
        if (fixedLanguage != null) {
          docLanguage = fixedLanguage;
        }
        currentLanguage = docLanguage;
      }
      lt = new WtLanguageTool(currentLanguage, config.getMotherTongue(),
          new UserConfig(config.getConfigurableValues(), linguServices), config, extraRemoteRules, 
          noLtSpeller, checkImpressDocument, testMode);
      config.initStyleCategories(lt.getAllRules());
      recheck = false;
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
        if (runTime > WtOfficeTools.TIME_TOLERANCE) {
          WtMessageHandler.printToLogFile("Time to init Language Tool: " + runTime);
        }
      }
      return lt;
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return lt;
  }

  /**
   * Initialize single documents, prepare text level rules and start queue
   */
  public void initDocuments(boolean resetCache) throws Throwable {
    long startTime = 0;
    if (debugModeTm) {
      startTime = System.currentTimeMillis();
    }
    setConfigValues(config, lt);
    if (debugModeTm) {
      long runTime = System.currentTimeMillis() - startTime;
      if (runTime > WtOfficeTools.TIME_TOLERANCE) {
        WtMessageHandler.printToLogFile("Time to init Documents: " + runTime);
      }
    }
  }
  
  /**
   * Get current locale language
   */
  public Locale getLocale() {
    return locale;
  }
  
  /**
   * Get list of single documents
   */
  public List<WtSingleDocument> getDocuments() {
    return documents;
  }

  /**
   * true, if LanguageTool is switched off
   */
  public boolean isBackgroundCheckOff() {
    return noBackgroundCheck;
  }

  /**
   * true, if Java look and feel is set
   */
  public static boolean isJavaLookAndFeelSet() {
    return javaLookAndFeelSet >= 0;
  }

  public static int getJavaLookAndFeelSet() {
    return javaLookAndFeelSet;
  }

  /**
   *  Toggle Switch Off / On of LT
   *  return true if toggle was done 
   */
  public boolean toggleNoBackgroundCheck() throws Throwable {
    if (docLanguage == null) {
      docLanguage = getCurrentLanguage();
    }
    if (config == null) {
      config = new WtConfiguration(configDir, configFile, docLanguage, true);
    }
    noBackgroundCheck = !noBackgroundCheck;
    setRecheck();
    config.saveNoBackgroundCheck(noBackgroundCheck, docLanguage);
    for (WtSingleDocument document : documents) {
      document.setConfigValues(config);
    }
    return true;
  }

  /**
   * Set use original spell und grammar dialog (for OO and old LO)
   */
  public void setUseOriginalCheckDialog() {
    useOrginalCheckDialog = true;
  }
  
  /**
   * Set use original spell und grammar dialog (for OO and old LO)
   */
  public boolean useOriginalCheckDialog() {
    return useOrginalCheckDialog;
  }
  
  /**
   * change configuration profile 
   */
  private void changeProfile(String profile) {
    try {
      if (profile == null) {
        profile = "";
      }
      WtMessageHandler.printToLogFile("change to profile: " + profile);
      String currentProfile = config.getCurrentProfile();
      if (currentProfile == null) {
        currentProfile = "";
      }
      if (profile.equals(currentProfile)) {
        WtMessageHandler.printToLogFile("profile == null or profile equals current profile: Not changed");
        return;
      }
      List<String> definedProfiles = config.getDefinedProfiles();
      if (!profile.isEmpty() && (definedProfiles == null || !definedProfiles.contains(profile))) {
        WtMessageHandler.showMessage("profile '" + profile + "' not found");
      } else {
        List<String> saveProfiles = new ArrayList<>();
        saveProfiles.addAll(config.getDefinedProfiles());
        config.initOptions();
        config.loadConfiguration(profile == null ? "" : profile);
        config.setCurrentProfile(profile);
        config.addProfiles(saveProfiles);
        config.saveConfiguration(getCurrentDocument().getLanguage());
        resetConfiguration();
      }
    } catch (IOException e) {
      WtMessageHandler.showError(e);
    }
  }
  
  /**
   * We leave spell checking to OpenOffice/LibreOffice.
   * @return false
   */
  public final boolean isSpellChecker() {
    return true;
  }
  
  /**
   * Returns extra remote rules
   */
  public List<Rule> getExtraRemoteRules() {
    return extraRemoteRules;
  }

  /**
   * Returns xContext
   */
  public XComponentContext getContext() {
    return xContext;
  }

  /**
   * Runs LT options dialog box.
   */
  public void runOptionsDialog() {
    try {
      WtConfiguration config = getConfiguration();
      Language lang = config.getDefaultLanguage();
      if (lang == null) {
        lang = getCurrentLanguage();
      }
      if (lang == null) {
        return;
      }
      WtLanguageTool lTool = lt;
      if (lTool == null || !lang.equals(docLanguage)) {
        docLanguage = lang;
        lTool = initLanguageTool();
        config = WtDocumentsHandler.config;
      }
      config.initStyleCategories(lTool.getAllRules());
      WtConfigThread configThread = new WtConfigThread(lang, config, lTool, this);
      configThread.start();
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }

  /**
   * @return An array of Locales supported by LT
   */
  public final static Locale[] getLocales() {
    try {
      List<Locale> locales = new ArrayList<>();
      Locale locale = null;
      for (Language lang : Languages.get()) {
        if (lang.getCountries().length == 0) {
          if (lang.getDefaultLanguageVariant() != null) {
            if (lang.getDefaultLanguageVariant().getVariant() != null) {
              locale = new Locale(lang.getDefaultLanguageVariant().getShortCode(),
                  lang.getDefaultLanguageVariant().getCountries()[0], lang.getDefaultLanguageVariant().getVariant());
            } else if (lang.getDefaultLanguageVariant().getCountries().length != 0) {
              locale = new Locale(lang.getDefaultLanguageVariant().getShortCode(),
                  lang.getDefaultLanguageVariant().getCountries()[0], "");
            } else {
              locale = new Locale(lang.getDefaultLanguageVariant().getShortCode(), "", "");
            }
          }
          else if (lang.getVariant() != null) {  // e.g. Esperanto
            locale =new Locale(LIBREOFFICE_SPECIAL_LANGUAGE_TAG, "", lang.getShortCodeWithCountryAndVariant());
          } else {
            locale = new Locale(lang.getShortCode(), "", "");
          }
          if (locales != null && !WtOfficeTools.containsLocale(locales, locale)) {
            locales.add(locale);
          }
        } else {
          for (String country : lang.getCountries()) {
            if (lang.getVariant() != null) {
              locale = new Locale(LIBREOFFICE_SPECIAL_LANGUAGE_TAG, country, lang.getShortCodeWithCountryAndVariant());
            } else {
              locale = new Locale(lang.getShortCode(), country, "");
            }
            if (locales != null && !WtOfficeTools.containsLocale(locales, locale)) {
              locales.add(locale);
            }
          }
        }
      }
      return locales == null ? new Locale[0] : locales.toArray(new Locale[0]);
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      return new Locale[0];
    }
  }

  /**
   * Add a listener that allow re-checking the document after changing the
   * options in the configuration dialog box.
   * 
   * @param eventListener the listener to be added
   * @return true if listener is non-null and has been added, false otherwise
   */
  public final boolean addLinguServiceEventListener(XLinguServiceEventListener eventListener) {
    if (eventListener == null) {
      return false;
    }
    xEventListeners.add(eventListener);
    return true;
  }

  /**
   * Remove a listener from the event listeners list.
   * 
   * @param eventListener the listener to be removed
   * @return true if listener is non-null and has been removed, false otherwise
   */
  public final boolean removeLinguServiceEventListener(XLinguServiceEventListener eventListener) {
    if (eventListener == null) {
      return false;
    }
    if (xEventListeners.contains(eventListener)) {
      xEventListeners.remove(eventListener);
      return true;
    }
    return false;
  }

  /**
   * Inform listener that the document should be rechecked for grammar and style check.
   */
  public boolean resetCheck() {
    return resetCheck(LinguServiceEventFlags.PROOFREAD_AGAIN);
  }

  /**
   * Inform listener that the doc should be rechecked for a special event flag.
   */
  public boolean resetCheck(short eventFlag) {
    if (!xEventListeners.isEmpty()) {
      for (XLinguServiceEventListener xEvLis : xEventListeners) {
        if (xEvLis != null) {
          LinguServiceEvent xEvent = new LinguServiceEvent();
          xEvent.nEvent = eventFlag;
          xEvLis.processLinguServiceEvent(xEvent);
        }
      }
      return true;
    }
    return false;
  }

  /**
   * Configuration has be changed
   */
  public void resetConfiguration() {
    linguServices = null;
    if (config != null) {
      noBackgroundCheck = config.noBackgroundCheck();
    }
    javaLookAndFeelSet = -1;
    resetDocument();
  }

  /**
   * Inform listener (grammar checking iterator) that options have changed and
   * the doc should be rechecked.
   */
  public void resetDocument() {
    setRecheck();
    resetCheck();
  }

  /**
   * Triggers the events from LT menu
   */
  public void trigger(String sEvent) {
    try {
      WtMessageHandler.printToLogFile("Trigger event: " + sEvent);
      if ("noAction".equals(sEvent)) {  //  special dummy action
        return;
      }
      if (getCurrentDocument() == null) {
        WtMessageHandler.printToLogFile("Trigger event: CurrentDocument == null");
        return;
      }
      long startTime = 0;
      if (debugModeTm) {
        startTime = System.currentTimeMillis();
      }
      if (!testDocLanguage(true)) {
        WtMessageHandler.printToLogFile("Test for document language failed: Can't trigger event: " + sEvent);
        return;
      }
      if ("configure".equals(sEvent)) {
        closeDialogs();
        runOptionsDialog();
      } else if ("about".equals(sEvent)) {
        if (aboutDialog != null) {
          aboutDialog.close();
          aboutDialog = null;
        }
        if (!isJavaLookAndFeelSet()) {
          setJavaLookAndFeel();
        }
        AboutDialogThread aboutThread = new AboutDialogThread(messages, xContext);
        aboutThread.start();
      } else if ("refreshCheck".equals(sEvent)) {
        resetDocument();
      } else if ("toggleNoBackgroundCheck".equals(sEvent) || "backgroundCheckOn".equals(sEvent) || "backgroundCheckOff".equals(sEvent)) {
        if (toggleNoBackgroundCheck()) {
          resetCheck();
        }
      } else if (sEvent.startsWith("profileChangeTo_")) {
        String profile = sEvent.substring(16);
        changeProfile(profile);
      } else if ("checkDialog".equals(sEvent) || "checkAgainDialog".equals(sEvent)) {
        if ("checkDialog".equals(sEvent) ) {
          WtOfficeTools.dispatchCmd(".uno:SpellingAndGrammarDialog", xContext);
        } else {
          WtOfficeTools.dispatchCmd(".uno:RecheckDocument", xContext);
        }
      } else if ("nextError".equals(sEvent)) {
        if (isBackgroundCheckOff()) {
          WtMessageHandler.showMessage(messages.getString("loExtSwitchOffMessage"));
          return;
        }
        WtOfficeTools.dispatchCmd(".uno:SpellingAndGrammarDialog", xContext);
      } else if ("remoteHint".equals(sEvent)) {
        if (getConfiguration().useOtherServer()) {
          WtMessageHandler.showMessage(MessageFormat.format(messages.getString("loRemoteInfoOtherServer"), 
              getConfiguration().getServerUrl()));
        } else {
          WtMessageHandler.showMessage(messages.getString("loRemoteInfoDefaultServer"));
        }
      } else {
        WtMessageHandler.printToLogFile("MultiDocumentsHandler: trigger: Sorry, don't know what to do, sEvent = " + sEvent);
      }
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
        if (runTime > WtOfficeTools.TIME_TOLERANCE) {
          WtMessageHandler.printToLogFile("Time to run trigger: " + runTime);
        }
      }
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
  }
  
  /**
   * Test the language of the document
   * switch the check to LT if possible and language is supported
   */
  boolean testDocLanguage(boolean showMessage) throws Throwable {
    try {
      if (docLanguage == null) {
        if (linguServices == null) {
          linguServices = getLinguisticServices();
        }
        if (xContext == null) {
          if (showMessage) { 
            WtMessageHandler.showMessage("There may be a installation problem! \nNo xContext!");
          }
          return false;
        }
        XComponent xComponent = WtOfficeTools.getCurrentComponent(xContext);
        if (xComponent == null) {
          if (showMessage) { 
            WtMessageHandler.showMessage("There may be a installation problem! \nNo xComponent!");
          }
          return false;
        }
        Locale locale;
        DocumentType docType = DocumentType.WRITER;
        locale = WtOfficeTools.getCursorLocale(xContext);
        try {
          int n = 0;
          while (locale == null && n < 100) {
            Thread.sleep(500);
            if (debugMode) {
              WtMessageHandler.printToLogFile("MultiDocumentsHandler: testDocLanguage: Try to get locale: n = " + n);
            }
            locale = WtOfficeTools.getCursorLocale(xContext);
            n++;
          }
        } catch (InterruptedException e) {
          WtMessageHandler.showError(e);
        }
        if (locale == null) {
          if (showMessage) {
            WtMessageHandler.showMessage("No Locale! LanguageTool can not be started!");
          } else {
            WtMessageHandler.printToLogFile("MultiDocumentsHandler: testDocLanguage: No Locale! LanguageTool can not be started!");
          }
          return false;
        } else if (!hasLocale(locale)) {
          String message = Tools.i18n(messages, "language_not_supported", locale.Language);
          WtMessageHandler.showMessage(message);
          return false;
        }
        if (debugMode) {
          WtMessageHandler.printToLogFile("MultiDocumentsHandler: testDocLanguage: locale: " + locale.Language + "-" + locale.Country);
        }
        if (!linguServices.setLtAsGrammarService(xContext, locale)) {
          if (showMessage) {
            WtMessageHandler.showMessage("Can not set LT as grammar check service! LanguageTool can not be started!");
          } else {
            WtMessageHandler.printToLogFile("MultiDocumentsHandler: testDocLanguage: Can not set LT as grammar check service! LanguageTool can not be started!");
          }
          return false;
        }
        if (docType != DocumentType.WRITER) {
          langForShortName = getLanguage(locale);
          if (langForShortName != null) {
            docLanguage = langForShortName;
            this.locale = locale;
            extraRemoteRules.clear();
            lt = initLanguageTool();
            initDocuments(true);
          }
          return true;
        } else {
          resetCheck();
          if (showMessage) {
            WtMessageHandler.showMessage(messages.getString("loNoGrammarCheckWarning"));
          }
          return false;
        }
      }
      return true;
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return false;
  }

  /**
   * Test if the needed java version is installed
   */
  public boolean javaVersionOkay() {
    String version = System.getProperty("java.version");
    if (version != null
        && (version.startsWith("1.0") || version.startsWith("1.1")
            || version.startsWith("1.2") || version.startsWith("1.3")
            || version.startsWith("1.4") || version.startsWith("1.5")
            || version.startsWith("1.6") || version.startsWith("1.7"))) {
      WtMessageHandler.showMessage("Error: LanguageTool requires Java 8 or later. Current version: " + version);
      return false;
    }
    return true;
  }
  
  /** Set Look and Feel for Java Swing Components
   * 
   */
  public static void setJavaLookAndFeel() {
    if (javaLookAndFeelSet < 0) {
      try {
        WtConfiguration config = WtDocumentsHandler.config;
        if (config == null) {
          Locale locale = new Locale("en","US","");
          config = getConfiguration(getLanguage(locale));
        }
        javaLookAndFeelSet = config == null ? -1 : config.getThemeSelection();
        WtGeneralTools.setJavaLookAndFeel(javaLookAndFeelSet < 0 ? 0 : javaLookAndFeelSet);
      } catch (Throwable t) {
        try {
          UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Throwable e) {
        }
      }
    }
  }
  
  /**
   * class to run the about dialog
   */
  private class AboutDialogThread extends Thread {

    private final ResourceBundle messages;
    private final XComponentContext xContext;

    AboutDialogThread(ResourceBundle messages, XComponentContext xContext) {
      this.messages = messages;
      this.xContext = xContext;
    }

    @Override
    public void run() {
      try {
        aboutDialog = new WtAboutDialog(messages);
        aboutDialog.show(xContext);
      } catch (Throwable t) {
        WtMessageHandler.showError(t);
      }
    }
    
  }

  /**
   * Called on rechecking the document - resets the ignore status for rules that
   * was set in the spelling dialog box or in the context menu.
   * 
   * The rules disabled in the config dialog box are left as intact.
   */
  public void resetIgnoreRules() {
    resetDisabledRules();
    setRecheck();
    docReset = true;
  }

  /**
   * Called when "Ignore" is selected e.g. in the context menu for an error.
   */
  public void ignoreRule(String ruleId, Locale locale) {
    addDisabledRule(WtOfficeTools.localeToString(locale), ruleId);
    setRecheck();
  }

  /**
   * Get the displayed service name for LT
   */
  public static String getServiceDisplayName(Locale locale) {
    return WtOfficeTools.WT_DISPLAY_SERVICE_NAME;
  }

  /**
   * remove internal stored text if document disposes
   */
  public void disposing(EventObject source) {
    //  the data of document will be removed by next call of getNumDocID
    //  to finish checking thread without crashing
    try {
      XComponent goneComponent = UnoRuntime.queryInterface(XComponent.class, source.Source);
      if (goneComponent == null) {
        WtMessageHandler.printToLogFile("MultiDocumentsHandler: disposing: xComponent of closed document is null");
      } else {
        setContextOfClosedDoc(goneComponent);
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }

  /**
   * class to run a dialog in a separate thread
   * closing if lost focus
   */
  public static class WaitDialogThread extends Thread {
    private final String dialogName;
    private final String text;
    private int max;
    private JDialog dialog = null;
    private boolean isCanceled = false;
    private JProgressBar progressBar;

    public WaitDialogThread(String dialogName, String text) {
      this.dialogName = dialogName;
      this.text = text;
      progressBar = new JProgressBar();
      if (!isJavaLookAndFeelSet()) {
        WtDocumentsHandler.setJavaLookAndFeel();
      }
    }

    @Override
    public void run() {
      try {
        JLabel textLabel = new JLabel(text);
        JButton cancelBottom = new JButton(messages.getString("guiCancelButton"));
        cancelBottom.addActionListener(e -> {
          close_intern();
        });
        progressBar.setIndeterminate(true);
        dialog = new JDialog();
        Container contentPane = dialog.getContentPane();
        dialog.setName("InformationThread");
        dialog.setTitle(dialogName);
        dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
        dialog.addWindowListener(new WindowListener() {
          @Override
          public void windowOpened(WindowEvent e) {
          }
          @Override
          public void windowClosing(WindowEvent e) {
            close_intern();
          }
          @Override
          public void windowClosed(WindowEvent e) {
          }
          @Override
          public void windowIconified(WindowEvent e) {
          }
          @Override
          public void windowDeiconified(WindowEvent e) {
          }
          @Override
          public void windowActivated(WindowEvent e) {
          }
          @Override
          public void windowDeactivated(WindowEvent e) {
          }
        });
        JPanel panel = new JPanel();
        panel.setLayout(new GridBagLayout());
        GridBagConstraints cons = new GridBagConstraints();
        cons.insets = new Insets(16, 24, 16, 24);
        cons.gridx = 0;
        cons.gridy = 0;
        cons.weightx = 1.0f;
        cons.weighty = 10.0f;
        cons.anchor = GridBagConstraints.CENTER;
        cons.fill = GridBagConstraints.BOTH;
        panel.add(textLabel, cons);
        cons.gridy++;
        panel.add(progressBar, cons);
        cons.gridy++;
        cons.fill = GridBagConstraints.NONE;
        panel.add(cancelBottom, cons);
        contentPane.setLayout(new GridBagLayout());
        cons = new GridBagConstraints();
        cons.insets = new Insets(16, 32, 16, 32);
        cons.gridx = 0;
        cons.gridy = 0;
        cons.weightx = 1.0f;
        cons.weighty = 1.0f;
        cons.anchor = GridBagConstraints.NORTHWEST;
        cons.fill = GridBagConstraints.BOTH;
        contentPane.add(panel);
        dialog.pack();
        Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
        Dimension frameSize = dialog.getSize();
        dialog.setLocation(screenSize.width / 2 - frameSize.width / 2,
            screenSize.height / 2 - frameSize.height / 2);
        dialog.setAutoRequestFocus(true);
        dialog.setAlwaysOnTop(true);
        dialog.toFront();
//        if (debugMode) {
          WtMessageHandler.printToLogFile(dialogName + ": run: Dialog is running");
//        }
        dialog.setVisible(true);
        if (isCanceled) {
          dialog.setVisible(false);
          dialog.dispose();
        }
      } catch (Throwable t) {
        WtMessageHandler.showError(t);
      }
    }
    
    public boolean canceled() {
      return isCanceled;
    }
    
    public void close() {
      close_intern();
    }
    
    private void close_intern() {
      try {
//        if (debugMode) {
          WtMessageHandler.printToLogFile("WaitDialogThread: close: Dialog closed");
//        }
        isCanceled = true;
        if (dialog != null) {
          if (dialog.isVisible()) {
            dialog.setVisible(false);
          }
          dialog.dispose();
        }
      } catch (Throwable t) {
        WtMessageHandler.showError(t);
      }
    }
    
    public void initializeProgressBar(int min, int max) {
      if (progressBar != null) {
        this.max = max;
        progressBar.setMinimum(min);
        progressBar.setMaximum(max);
        progressBar.setStringPainted(true);
        progressBar.setIndeterminate(false);
      }
    }
    
    public void setValueForProgressBar(int val, boolean setText) {
      if (progressBar != null) {
        progressBar.setValue(val);
        progressBar.setIndeterminate(false);
        if (setText) {
          int p = (int) (((val * 100) / max) + 0.5);
          progressBar.setString(p + " %  ( " + val + " / " + max + " )");
          progressBar.setStringPainted(true);
        }
      }
    }
    
    public void setIndeterminate() {
      progressBar.setIndeterminate(true);
    }
  }
}


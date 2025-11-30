/**
 * Translates Chinese text to English.
 *
 * @param {string} text The text to translate.
 * @return {string} The translated text.
 * @customfunction
 */
function translateChineseToEnglish(text) {
  if (typeof text !== 'string' || text.length === 0) {
    return "Input text cannot be empty.";
  }

  // Detect if the text contains Chinese characters.
  // This is a simple check and might not be 100% accurate.
  const containsChinese = (str) => {
    return /[\u4e00-\u9fa5]/.test(str);
  };

  if (!containsChinese(text)) {
    return text; // Return original text if no Chinese characters are found.
  }

  try {
    const translatedText = LanguageApp.translate(text, 'zh', 'en');
    return translatedText.toLowerCase();
  } catch (e) {
    return "Error in translation: " + e.toString();
  }
}

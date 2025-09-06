// functions/legal-query/index.js
import { Client, Databases, Query, Storage, ID } from 'node-appwrite';
import stringSimilarity from 'string-similarity';
import { config } from 'dotenv';
import XLSX from 'xlsx';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  TextDirection,
} from 'docx';
import {
  GoogleGenerativeAI,
  HarmCategory,
  HarmBlockThreshold,
} from '@google/generative-ai';
import {
  law_category,
  court_ruling_subjects,
  form_title,
  lawterms,
} from './full_data.js';

config({ debug: false });

// Utility functions (unchanged from your code)
const scoreLawRelevance = (law, tagWeights, queryTerms, selectedCategories) => {
  let tagScore = 0;
  let queryScore = 0;
  const content = (law.content || '').toLowerCase();
  const subcategory = (law.subcategory || '').toLowerCase();

  Object.entries(tagWeights).forEach(([tag, weight]) => {
    if (content.includes(tag.toLowerCase())) {
      tagScore += weight;
    }
  });

  queryTerms.forEach((term) => {
    if (
      content.includes(term.toLowerCase()) ||
      subcategory.includes(term.toLowerCase())
    ) {
      queryScore += 1;
    }
  });

  const categoryMatch = selectedCategories.includes(law.subcategory) ? 10 : 0;
  return tagScore + queryScore + categoryMatch;
};

const getAppwriteFile = async (bucketId, fileId, formTitle, url) => {
  try {
    console.log(`Fetching file ${fileId} for ${formTitle}`);
    console.log(`Using URL: ${url}`);

    const startTime = Date.now();
    const response = await fetch(url, {
      headers: {
        'X-Appwrite-Project': process.env.APPWRITE_PROJECT_ID,
        'X-Appwrite-Key': process.env.APPWRITE_API_KEY,
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const fileData = Buffer.from(await response.arrayBuffer());
    const duration = Date.now() - startTime;
    console.log(
      `Fetched file in ${duration}ms, size: ${fileData.length} bytes`
    );

    if (!Buffer.isBuffer(fileData) || fileData.length === 0) {
      throw new Error(`Invalid file data: ${fileData.constructor.name}`);
    }

    if (fileData.length > 50 * 1024 * 1024) {
      throw new Error(`File exceeds 50MB limit: ${fileData.length} bytes`);
    }

    const mimeType = 'application/msword';
    console.log(`Using MIME type: ${mimeType}`);
    return { fileData, mimeType };
  } catch (e) {
    console.error(`File fetch failed: ${e.message}`);
    throw e;
  }
};

const createWordDocument = async (content) => {
  try {
    console.log('Creating Word document with enhanced RTL settings');

    const doc = new Document({
      compatibility: {
        doNotUseEastAsianBreakRules: true,
        useNormalStyleForList: true,
      },
      styles: {
        default: {
          paragraph: {
            bidi: true,
            alignment: AlignmentType.RIGHT,
            textDirection: TextDirection.RIGHT_TO_LEFT,
            run: {
              font: 'B Nazanin',
              size: 24,
            },
          },
        },
      },
      sections: [
        {
          properties: {
            bidi: true,
          },
          children: [
            new Paragraph({
              style: 'default',
              alignment: AlignmentType.CENTER,
              textDirection: TextDirection.RIGHT_TO_LEFT,
              children: [
                new TextRun({
                  text: 'پاسخ مشاوره حقوقی',
                  bold: true,
                  size: 36,
                  font: 'B Nazanin',
                }),
              ],
            }),
            new Paragraph({
              style: 'default',
              alignment: AlignmentType.CENTER,
              textDirection: TextDirection.RIGHT_TO_LEFT,
              children: [
                new TextRun({
                  text: '――――――――――――――――――――',
                  color: '808080',
                }),
              ],
              spacing: { after: 200 },
            }),
            ...content.split('\n\n').map(
              (paragraph) =>
                new Paragraph({
                  style: 'default',
                  alignment: AlignmentType.RIGHT,
                  textDirection: TextDirection.RIGHT_TO_LEFT,
                  children: [
                    new TextRun({
                      text: paragraph,
                      font: 'B Nazanin',
                      size: 24,
                    }),
                  ],
                  spacing: { before: 100, after: 100 },
                })
            ),
            new Paragraph({
              style: 'default',
              alignment: AlignmentType.CENTER,
              textDirection: TextDirection.RIGHT_TO_LEFT,
              children: [
                new TextRun({
                  text: 'مشاوره حقوقی توسط ربات وکیل جی‌بی‌ام',
                  italics: true,
                  color: '808080',
                  size: 20,
                }),
              ],
              spacing: { before: 300 },
            }),
          ],
        },
      ],
    });

    const buffer = await Packer.toBuffer(doc);
    console.log(`Generated Word document, size: ${buffer.length} bytes`);
    return buffer;
  } catch (e) {
    console.error(`Word document creation failed: ${e.message}`);
    throw e;
  }
};

const batchQueries = async (databases, queries, collection, limit = 100) => {
  try {
    const batches = [];
    for (let i = 0; i < queries.length; i += 10) {
      batches.push(queries.slice(i, i + 10));
    }

    const results = [];
    for (const batch of batches) {
      let offset = 0;
      let totalFetched = 0;
      while (true) {
        const batchResult = await databases.listDocuments(
          process.env.APPWRITE_DATABASE_ID,
          collection,
          [...batch, Query.limit(limit), Query.offset(offset)],
          limit
        );
        results.push(batchResult);
        totalFetched += batchResult.documents.length;

        console.log(
          `[${collection}] Fetched ${batchResult.documents.length} documents at offset ${offset}, total: ${totalFetched}`
        );

        if (batchResult.documents.length < limit) break;
        offset += limit;
      }
    }
    console.log(
      `[${collection}] Total documents fetched: ${results.flatMap((r) => r.documents).length}`
    );
    return results;
  } catch (e) {
    console.error(`[${collection}] Batch query failed: ${e.message}`);
    return [];
  }
};

const getGenerativeModel = async (
  prompt,
  stepName,
  modelPreference = 'gemini-2.0-flash',
  step = 0,
  retries = 3,
  baseDelay = 1000
) => {
  const geminiKeys = [
    process.env.GEMINI_API_KEY,
    process.env.GEMINI_API_KEY_2,
    process.env.GEMINI_API_KEY_3,
    process.env.GEMINI_API_KEY_4,
    process.env.GEMINI_API_KEY_5,
    process.env.GEMINI_API_KEY_6,
    process.env.GEMINI_API_KEY_7,
  ].filter(Boolean);

  const modelVariants = [
    'gemini-2.0-flash',
    'gemini-2.5-flash',
    'gemini-1.5-flash',
  ];

  const startingKeyIndex = step % geminiKeys.length;
  const keysToTry = [
    ...geminiKeys.slice(startingKeyIndex),
    ...geminiKeys.slice(0, startingKeyIndex),
  ];

  console.log(
    `[${stepName}] Starting with key ${keysToTry[0].slice(-4)} for step ${step}`
  );

  let lastError = null;

  for (const apiKey of keysToTry) {
    for (let attempt = 0; attempt < retries; attempt++) {
      try {
        console.log(
          `[${stepName}] Attempt ${attempt + 1}/${retries} with model ${modelPreference} and key ${apiKey.slice(-4)}`
        );

        const genAI = new GoogleGenerativeAI(apiKey);
        const modelConfig = {
          model: modelPreference,
          safetySettings: [
            {
              category: HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT,
              threshold: HarmBlockThreshold.BLOCK_NONE,
            },
          ],
        };
        if (step === 3 || step === 4) {
          modelConfig.tools = [{ googleSearch: {} }];
        }
        const model = genAI.getGenerativeModel(modelConfig);

        const result = await model.generateContent(prompt);
        const response = await result.response;
        const text = response.text();

        console.log(
          `[${stepName}] Success with model ${modelPreference} and key ${apiKey.slice(-4)}`
        );
        console.log(`[${stepName}] Response length: ${text.length} chars`);

        return {
          text,
          model: modelPreference,
          apiKey: apiKey.slice(-4),
        };
      } catch (e) {
        lastError = e;
        console.error(
          `[${stepName}] Attempt ${attempt + 1} failed: ${e.message}`
        );

        if (e.message.includes('429 Too Many Requests')) {
          const retryDelayMatch = e.message.match(/retryDelay":"(\d+)s"/);
          const retryDelay = retryDelayMatch
            ? parseInt(retryDelayMatch[1]) * 1000
            : baseDelay * (attempt + 1);

          console.log(
            `[${stepName}] Rate limit hit for key ${apiKey.slice(-4)}, moving to next key`
          );
          break;
        } else if (e.message.includes('503 Service Unavailable')) {
          console.log(
            `[${stepName}] Model ${modelPreference} overloaded for key ${apiKey.slice(-4)}, moving to next key`
          );
          break;
        } else if (e.message.includes('404 Not Found')) {
          console.log(
            `[${stepName}] Model ${modelPreference} not found for key ${apiKey.slice(-4)}, moving to next key`
          );
          break;
        } else {
          await new Promise((res) =>
            setTimeout(res, baseDelay * (attempt + 1))
          );
        }
      }
    }
  }

  console.log(
    `[${stepName}] Exhausted all keys for ${modelPreference}, trying other models`
  );

  for (const apiKey of keysToTry) {
    const fallbackModels = modelVariants.filter((m) => m !== modelPreference);
    for (const modelName of fallbackModels) {
      for (let attempt = 0; attempt < retries; attempt++) {
        try {
          console.log(
            `[${stepName}] Attempt ${attempt + 1}/${retries} with model ${modelName} and key ${apiKey.slice(-4)}`
          );

          const genAI = new GoogleGenerativeAI(apiKey);
          const modelConfig = {
            model: modelName,
            safetySettings: [
              {
                category: HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT,
                threshold: HarmBlockThreshold.BLOCK_NONE,
              },
            ],
          };
          if (step === 4) {
            modelConfig.tools = [{ googleSearch: {} }];
          }
          const model = genAI.getGenerativeModel(modelConfig);

          const result = await model.generateContent(prompt);
          const response = await result.response;
          const text = response.text();

          console.log(
            `[${stepName}] Success with model ${modelName} and key ${apiKey.slice(-4)}`
          );
          console.log(`[${stepName}] Response length: ${text.length} chars`);

          return {
            text,
            model: modelName,
            apiKey: apiKey.slice(-4),
          };
        } catch (e) {
          lastError = e;
          console.error(
            `[${stepName}] Attempt ${attempt + 1} failed: ${e.message}`
          );

          if (e.message.includes('429 Too Many Requests')) {
            const retryDelayMatch = e.message.match(/retryDelay":"(\d+)s"/);
            const retryDelay = retryDelayMatch
              ? parseInt(retryDelayMatch[1]) * 1000
              : baseDelay * (attempt + 1);

            console.log(
              `[${stepName}] Rate limit hit for key ${apiKey.slice(-4)}, moving to next key`
            );
            break;
          } else if (e.message.includes('503 Service Unavailable')) {
            console.log(
              `[${stepName}] Model ${modelName} overloaded for key ${apiKey.slice(-4)}, moving to next key`
            );
            break;
          } else if (e.message.includes('404 Not Found')) {
            console.log(
              `[${stepName}] Model ${modelName} not found for key ${apiKey.slice(-4)}, skipping model`
            );
            break;
          } else {
            await new Promise((res) =>
              setTimeout(res, baseDelay * (attempt + 1))
            );
          }
        }
      }
    }
  }

  console.error(
    `[${stepName}] All attempts failed: ${lastError?.message || 'Unknown error'}`
  );
  return {
    text: 'متأسفیم، در تولید پاسخ مشکلی پیش آمد. لطفاً بعداً دوباره تلاش کنید.',
    model: modelPreference || 'none',
    apiKey: 'none',
  };
};

const parseAIResponse = (text, field) => {
  try {
    let jsonString = text.trim();
    console.log(`[${field}] Raw response: ${jsonString.substring(0, 200)}...`);

    if (jsonString.startsWith('```json') && jsonString.endsWith('```')) {
      jsonString = jsonString.slice(7, -3).trim();
    } else if (jsonString.startsWith('```') && jsonString.endsWith('```')) {
      jsonString = jsonString.slice(3, -3).trim();
    }

    const parsed = JSON.parse(jsonString);
    console.log(`[${field}] Parsed successfully`);
    return parsed;
  } catch (e) {
    console.error(`[${field}] Parse failed: ${e.message}`);
    return { error: `Parse error: ${e.message}` };
  }
};

// processLegalQuery (unchanged)
const processLegalQuery = async (userQuery, conversation, databases) => {
  console.log('==================================');
  console.log('PROCESSING LEGAL QUERY');
  console.log('==================================');
  console.log(`User query: ${userQuery.substring(0, 50)}...`);
  console.log(`Conversation length: ${conversation.length} chars`);
  console.log('Databases object:', databases);
  console.log(
    'APPWRITE_DATABASE_ID:',
    process.env.APPWRITE_DATABASE_ID || 'Not defined'
  );

  if (!process.env.APPWRITE_DATABASE_ID) {
    console.error('Environment variable APPWRITE_DATABASE_ID is not defined');
    return {
      is_meaningful: false,
      error: 'Configuration error: APPWRITE_DATABASE_ID is not defined',
      step4Response: 'متأسفیم، خطایی در پیکربندی رخ داد.',
      laws: [],
      courtRulingDocuments: [],
      matchedTerms: [],
      formDocs: [],
      customForm: null,
    };
  }

  let lawsText = 'هیچ قانونی یافت نشد';
  let courtRulingsText = 'هیچ حکمی یافت نشد';
  let formsText = 'هیچ فرمی یافت نشد';
  let terminologyText = 'هیچ اصطلاحی یافت نشد';
  let finalLaws = [];
  let courtRulingDocuments = [];
  let matchedTerms = [];
  let formDocs = [];
  let customForm = null;

  try {
    console.log('----------------------------------------');
    console.log('STEP 0: CONVERSATION VALIDATION');
    console.log('----------------------------------------');

    const validationPrompt = `
You are a Persian lawyer AI agent. Analyze the following conversation and user query from a chat session. Determine if the conversation or query is meaningful and relevant to a legal query requiring assistance from an AI lawyer agent. Return **only** a JSON object with a single boolean field:

{
  "is_meaningful": boolean
}

Conversation:
${conversation}

User Query:
${userQuery}

- Return "is_meaningful": true if the conversation or query contains at least one message that appears to be a genuine legal query related to Iranian laws, even if it's a single message or lacks extensive context.
- Return "is_meaningful": false only if both the conversation and query are clearly irrelevant, malicious, or unrelated to legal assistance (e.g., spam, greetings, or non-legal topics).
- Examples of valid legal queries include questions about property, contracts, permits, legal procedures, or disputes, even if brief or informal.
`;

    let isMeaningful = false;
    try {
      console.log('Sending validation prompt to AI...');
      console.log(`Conversation content: ${conversation.substring(0, 100)}...`);
      const validationResponse = await getGenerativeModel(
        validationPrompt,
        'Step 0: Validation',
        'gemini-2.0-flash',
        0
      );

      const validationResult = parseAIResponse(
        validationResponse.text,
        'Step 0 Response'
      );

      isMeaningful = validationResult.is_meaningful || false;
      console.log(`Validation result: ${isMeaningful}`);
    } catch (e) {
      console.error('Step 0 failed:', e.message);
      return { is_meaningful: false, error: 'Validation failed' };
    }

    if (!isMeaningful) {
      console.log('Conversation not meaningful - ending processing');
      return { is_meaningful: false };
    }

    console.log('----------------------------------------');
    console.log('STEP 1: LAW CATEGORIES AND TAGS');
    console.log('----------------------------------------');

    const step1StaticPrompt = `
You are a component of a Persian lawyer AI agent that processes queries about Iranian laws. Your task is to analyze the user's Persian query and conversation history, and return **only** a valid JSON object in the exact format specified below. Do not include any additional text, explanations, or comments outside the JSON.

Given the query: "${userQuery}"

Conversation history (last 4 messages):
${conversation}

You have access to:
- Law categories (~${law_category.length} total): [LAW_CATEGORIES]

Return **only** this JSON structure:
{
  "laws": {
    "categories": ["string", ...],
    "tags": {"string": number, ...}
  }
}

- Select 5–15 law categories that closely match the query's legal context, considering the conversation history for additional context.
- Generate highly accurate legal tags for laws, based on the query and conversation. Aim for around 20 tags, but any number is acceptable. Prioritize precise legal terms (e.g., "ثبت شرکت", "پروانه کسب", "مالیات", "قانون تجارت", "تهران"). Avoid generic terms unless legally relevant. Assign each tag an importance weight (0–1, float) reflecting its relevance to the query and conversation.
- Ensure tags are in Persian, specific to Iranian laws, and aligned with the query's intent.
`;

    const step1Prompt = step1StaticPrompt.replace(
      '[LAW_CATEGORIES]',
      JSON.stringify(law_category)
    );

    let structuredFilters = { laws: { categories: [], tags: {} } };
    try {
      const step1Response = await getGenerativeModel(
        step1Prompt,
        'Step 1: Law Filters',
        'gemini-2.0-flash',
        1
      );

      structuredFilters = parseAIResponse(
        step1Response.text,
        'Step 1 Response'
      );

      console.log(
        `Generated ${structuredFilters.laws.categories.length} categories`
      );
      console.log(
        `Generated ${Object.keys(structuredFilters.laws.tags).length} tags`
      );
      console.log(
        'Sample categories:',
        structuredFilters.laws.categories.slice(0, 3)
      );
      console.log(
        'Sample tags:',
        Object.entries(structuredFilters.laws.tags).slice(0, 3)
      );
    } catch (e) {
      console.error('Step 1 failed:', e.message);
      structuredFilters = { laws: { categories: [], tags: {} } };
    }

    console.log('----------------------------------------');
    console.log('PRE-FILTERING COURT RULING TITLES');
    console.log('----------------------------------------');

    let filteredCourtRulingTitles = [];
    if (Object.keys(structuredFilters.laws.tags).length > 0) {
      const tags = Object.keys(structuredFilters.laws.tags);
      console.log('Tags for filtering court rulings:', tags);
      tags.forEach((tag) => {
        const matches = court_ruling_subjects.filter((title) =>
          title.includes(tag)
        );
        console.log(
          `Tag "${tag}" matched ${matches.length} court ruling titles (showing up to 3):`,
          matches.slice(0, 3)
        );
        filteredCourtRulingTitles.push(...matches);
      });
      filteredCourtRulingTitles = [...new Set(filteredCourtRulingTitles)];
      console.log(
        `Filtered court ruling titles to ${filteredCourtRulingTitles.length} based on tags`
      );
      console.log(
        'Sample filtered court ruling titles:',
        filteredCourtRulingTitles.slice(0, 3)
      );
    } else {
      console.log('No tags available, using all court ruling titles');
      filteredCourtRulingTitles = court_ruling_subjects;
    }

    console.log('----------------------------------------');
    console.log('STEP 2: COURT RULING TITLES');
    console.log('----------------------------------------');

    const step2StaticPrompt = `
You are a component of a Persian lawyer AI agent that processes queries about Iranian laws. Your task is to analyze the user's Persian query and conversation history, and return **only** a valid JSON object in the exact format specified below. Do not include any additional text, explanations, or comments outside the JSON.

Given the query: "${userQuery}"

Conversation history (last 4 messages):
${conversation}

You have access to:
- Court ruling titles (~${filteredCourtRulingTitles.length} total): [COURT_RULING_TITLES]

Return **only** this JSON structure:
{
  "court_ruling": {
    "titles": ["string", ...]
  }
}

- Select up to 20 unique titles from the provided court ruling titles array that closely match the query and its legal intent, considering the conversation history for additional context.
- Avoid duplicates. If no titles are highly relevant, select the closest matches or return an empty array if none are relevant.
- Ensure the selected titles align with the query's intent and are specific to Iranian laws.
`;

    const step2Prompt = step2StaticPrompt.replace(
      '[COURT_RULING_TITLES]',
      JSON.stringify(filteredCourtRulingTitles)
    );

    let courtRulingFilters = { court_ruling: { titles: [] } };
    try {
      const step2Response = await getGenerativeModel(
        step2Prompt,
        'Step 2: Court Rulings',
        'gemini-2.0-flash',
        2
      );

      courtRulingFilters = parseAIResponse(
        step2Response.text,
        'Step 2 Response'
      );

      console.log(
        `Selected ${courtRulingFilters.court_ruling.titles.length} court ruling titles`
      );
      console.log(
        'Selected court ruling titles:',
        courtRulingFilters.court_ruling.titles.slice(0, 10)
      );
    } catch (e) {
      console.error('Step 2 failed:', e.message);
      courtRulingFilters = { court_ruling: { titles: [] } };
    }

    console.log('----------------------------------------');
    console.log('RETRIEVING COURT RULINGS FROM DATABASE');
    console.log('----------------------------------------');

    const normalizeTitle = (title) => {
      if (!title || typeof title !== 'string' || title.trim() === '') {
        console.log(`Invalid title detected: ${JSON.stringify(title)}`);
        return null;
      }
      return title
        .trim()
        .replace(/\s+/g, ' ')
        .replace(/ي/g, 'ی')
        .replace(/ك/g, 'ک')
        .normalize('NFC');
    };

    if (courtRulingFilters.court_ruling.titles.length > 0) {
      try {
        const selectedTitles = [
          ...new Set(courtRulingFilters.court_ruling.titles),
        ].slice(0, 20);
        console.log('Selected titles for court ruling query:', selectedTitles);

        const normalizedTitles = selectedTitles
          .map(normalizeTitle)
          .filter((title) => title !== null)
          .map((title) => ({ query: Query.equal('title', title), title }));

        console.log(
          'Normalized court ruling titles:',
          normalizedTitles.map((t) => t.title).slice(1, 10)
        );

        courtRulingDocuments = [];
        for (let i = 0; i < normalizedTitles.length; i++) {
          const { query, title } = normalizedTitles[i];
          try {
            console.log(`Executing query ${i} for title: ${title}`);
            const response = await databases.listDocuments(
              process.env.APPWRITE_DATABASE_ID,
              'court_ruling',
              [query]
            );
            console.log(
              `Query ${i} for title "${title}" returned ${response.documents.length} documents`
            );
            courtRulingDocuments.push(...response.documents);
          } catch (e) {
            console.error(
              `Query ${i} for title "${title}" failed: ${e.message}`
            );
          }
        }

        courtRulingDocuments = Array.from(
          new Map(courtRulingDocuments.map((doc) => [doc.$id, doc])).values()
        ).slice(0, 20);

        console.log(
          `Found ${courtRulingDocuments.length} court ruling documents after deduplication`
        );
        console.log(
          'Retrieved court ruling titles:',
          courtRulingDocuments.map((doc) => doc.title).slice(0, 3)
        );
      } catch (e) {
        console.error('Court ruling retrieval failed:', e.message);
        courtRulingDocuments = [];
      }
    } else {
      console.log('No court ruling titles to query');
    }

    console.log('----------------------------------------');
    console.log('RETRIEVING LAWS FROM DATABASE');
    console.log('----------------------------------------');

    const maxLaws = 50;
    let allLaws = [];

    if (
      structuredFilters.laws.categories.length > 0 ||
      Object.keys(structuredFilters.laws.tags).length > 0
    ) {
      try {
        console.log(
          `Querying laws with ${structuredFilters.laws.categories.length} categories and ${Object.keys(structuredFilters.laws.tags).length} tags...`
        );

        const lawsQueries = [];
        if (structuredFilters.laws.categories.length > 0) {
          lawsQueries.push(
            Query.equal('subcategory', structuredFilters.laws.categories)
          );
        }
        if (Object.keys(structuredFilters.laws.tags).length > 0) {
          lawsQueries.push(
            Query.contains('content', Object.keys(structuredFilters.laws.tags))
          );
        }

        const lawsBatches = await batchQueries(databases, lawsQueries, 'laws');
        allLaws = lawsBatches
          .flatMap((batch) => batch.documents)
          .filter((doc) => doc);

        console.log(`Found ${allLaws.length} law documents`);
      } catch (e) {
        console.error('Law retrieval failed:', e.message);
        allLaws = [];
      }
    } else {
      console.log('No law categories or tags to query');
    }

    console.log('----------------------------------------');
    console.log('FILTERING AND SCORING LAWS');
    console.log('----------------------------------------');

    const lawsByCategory = {};
    allLaws.forEach((law) => {
      const category = law.subcategory || 'Uncategorized';
      if (!lawsByCategory[category]) lawsByCategory[category] = [];
      lawsByCategory[category].push(law);
    });

    finalLaws = [];
    Object.keys(lawsByCategory).forEach((category) => {
      const laws = lawsByCategory[category];
      if (laws.length <= 5) {
        finalLaws.push(...laws);
      } else {
        const scoredLaws = laws.map((law) => ({
          law,
          score: scoreLawRelevance(
            law,
            structuredFilters.laws.tags,
            userQuery.trim().split(/\s+/),
            structuredFilters.laws.categories
          ),
        }));
        scoredLaws.sort((a, b) => b.score - a.score);
        finalLaws.push(...scoredLaws.slice(0, 5).map((item) => item.law));
      }
    });
    finalLaws = finalLaws.slice(0, maxLaws);

    console.log(`Final laws after filtering: ${finalLaws.length}`);

    console.log('----------------------------------------');
    console.log('PRE-FILTERING FORM TITLES');
    console.log('----------------------------------------');

    let filteredFormTitles = [];
    if (Object.keys(structuredFilters.laws.tags).length > 0) {
      const tags = Object.keys(structuredFilters.laws.tags);
      console.log('Tags for filtering forms:', tags);
      tags.forEach((tag) => {
        const matches = form_title
          .filter((form) => form.title.includes(tag))
          .map((form) => form.title);
        console.log(
          `Tag "${tag}" matched ${matches.length} form titles (showing up to 3):`,
          matches.slice(0, 3)
        );
        filteredFormTitles.push(...matches);
      });
      filteredFormTitles = [...new Set(filteredFormTitles)];
      console.log(
        `Filtered form titles to ${filteredFormTitles.length} based on tags`
      );
      console.log(
        'Sample filtered form titles:',
        filteredFormTitles.slice(0, 3)
      );
    } else {
      console.log('No tags available, using all form titles');
      filteredFormTitles = form_title.map((form) => form.title);
    }

    console.log('----------------------------------------');
    console.log('STEP 3: TERMINOLOGIES, FORM TITLES, AND CUSTOM FORM');
    console.log('----------------------------------------');

    const step3StaticPrompt = `
You are a Persian lawyer AI agent that processes queries about Iranian laws. Your task is to analyze the user's query, conversation history, and the provided legal documents (laws and court rulings), and understand the query's intent. Return a JSON object with:
1. An array of up to 100 legal terminologies (in Persian) found in the documents that are ambiguous or require clarification.
2. An array of up to 25 titles from the provided form titles array that are relevant to the query.
3. If the user specifically requests a particular legal form (e.g., "مبایعه نامه", "دادخواست طلاق", "قرارداد کاری"), generate a custom form with a title and content tailored to the user's request, ensuring the form is legally compliant with Iranian laws and informed by the provided laws and court rulings.

Given the query: "${userQuery}"
Conversation history (last 4 messages):
${conversation}
Laws documents (up to ${maxLaws}): 
${lawsText}
Court ruling documents (up to 20):
${courtRulingsText}
Form titles (~${filteredFormTitles.length} total): 
${JSON.stringify(filteredFormTitles)}

**Instructions for Custom Form**:
- search internet to make sure our custom form is compliant with the latest legal requirements 
- If the user requests a specific legal form, create a detailed form that is legally accurate under Iranian law.
- Use the provided laws and court rulings to ensure the form's content is compliant and relevant. Reference specific laws (e.g., "ماده ۱۱۳۰ قانون مدنی") or court rulings (e.g., "رای شماره ۹۸۸ دیوان عالی") where applicable to justify the form's structure or clauses.
- If the provided laws or rulings are insufficient, indicate that additional legal references may be needed but provide the best possible form based on available data.
- The form should include placeholders for user-specific details (e.g., [نام طرفین], [تاریخ], [مبلغ]) and clear instructions for completion.

Return **only** this JSON structure:
{
  "unknown_terminologies": ["string", ...],
  "form_titles": ["string", ...],
  "custom_form": {
    "title": "string",
    "content": "string"
  }
}
`;

    let step3Response = {
      unknown_terminologies: [],
      form_titles: [],
      custom_form: null,
    };
    try {
      console.log(
        'Sending terminologies, form titles, and custom form prompt to AI...'
      );
      const step3Result = await getGenerativeModel(
        step3StaticPrompt,
        'Step 3: Terminologies, Form Titles, and Custom Form',
        'gemini-2.0-flash',
        3
      );

      step3Response = parseAIResponse(step3Result.text, 'Step 3 Response');

      console.log(
        `Found ${step3Response.unknown_terminologies.length} terminologies`
      );
      console.log(`Found ${step3Response.form_titles.length} form titles`);
      console.log(
        'Selected form titles:',
        step3Response.form_titles.slice(0, 10)
      );
      if (step3Response.custom_form) {
        console.log('Custom form generated:', step3Response.custom_form.title);
      }
    } catch (e) {
      console.error('Step 3 failed:', e.message);
      step3Response = {
        unknown_terminologies: [],
        form_titles: [],
        custom_form: null,
      };
    }

    console.log('----------------------------------------');
    console.log('MATCHING TERMS');
    console.log('----------------------------------------');

    matchedTerms = lawterms.filter((termObj) =>
      step3Response.unknown_terminologies.some((term) => term === termObj.term)
    );

    console.log(`Matched ${matchedTerms.length} legal terms`);

    console.log('----------------------------------------');
    console.log('RETRIEVING FORMS FROM DATABASE');
    console.log('----------------------------------------');

    if (step3Response.form_titles.length > 0) {
      try {
        const selectedFormTitles = [...new Set(step3Response.form_titles)];
        console.log('Selected titles for form query:', selectedFormTitles);

        const selectedIds = form_title
          .filter((form) => selectedFormTitles.includes(form.title))
          .map((form) => form.$id);

        console.log('Mapped $ids for form query:', selectedIds);

        if (selectedIds.length === 0) {
          console.log('No matching $ids found for selected titles');
        } else {
          const formQueries = [Query.equal('$id', selectedIds)];
          const formBatches = await batchQueries(
            databases,
            formQueries,
            'forms',
            100
          );
          formDocs = formBatches
            .flatMap((batch) => batch.documents)
            .filter((doc) => doc)
            .slice(0, 5);

          console.log(`Found ${formDocs.length} form documents`);
          console.log(
            'Retrieved form titles:',
            formDocs.map((doc) => doc.title).slice(0, 3)
          );
        }

        if (formDocs.length === 0) {
          console.log('No exact matches found, attempting fuzzy matching...');
          const allFormDocs = await databases.listDocuments(
            process.env.APPWRITE_DATABASE_ID,
            'forms',
            [Query.limit(100)]
          );
          formDocs = allFormDocs.documents
            .map((doc) => {
              let matchInfo;
              try {
                const bestMatch = stringSimilarity.findBestMatch(
                  doc.title,
                  selectedFormTitles
                );
                matchInfo = {
                  doc,
                  similarity: bestMatch.bestMatch.rating,
                  bestMatchTitle: bestMatch.bestMatch.target,
                };
              } catch (e) {
                console.warn(
                  'string-similarity failed, using fallback:',
                  e.message
                );
                const bestMatchTitle = selectedFormTitles.find((title) =>
                  doc.title.includes(title)
                );
                matchInfo = {
                  doc,
                  similarity: bestMatchTitle ? 'N/A (fallback)' : 'None',
                  bestMatchTitle: bestMatchTitle || 'No close match',
                };
              }
              return matchInfo;
            })
            .filter(
              (item) =>
                item.bestMatchTitle !== 'No close match' &&
                (item.similarity > 0.8 || item.similarity === 'N/A (fallback)')
            )
            .map((item) => {
              console.log(
                `Fuzzy matched form: DB Title: ${item.doc.title}, Best Match: ${item.bestMatchTitle}, Similarity: ${item.similarity}`
              );
              return item.doc;
            })
            .slice(0, 5);
          console.log(`Found ${formDocs.length} fuzzy matched form documents`);
        }
      } catch (e) {
        console.error('Form retrieval failed:', e.message);
        formDocs = [];
      }
    } else {
      console.log('No form titles to query');
    }

    if (step3Response.custom_form) {
      customForm = step3Response.custom_form;
    }

    console.log('----------------------------------------');
    console.log('STEP 4: FINAL LEGAL ADVICE');
    console.log('----------------------------------------');

    lawsText =
      finalLaws
        .map(
          (law) =>
            `Title: ${law.title || 'No title'}\nContent: ${law.content || 'No content'}`
        )
        .join('\n\n') || 'هیچ قانونی یافت نشد';

    courtRulingsText =
      courtRulingDocuments
        .slice(0, 20)
        .map(
          (ruling) =>
            `Title: ${ruling.title || 'No title'}\nRuling Group: ${ruling.ruling_group || 'N/A'}\nText: ${ruling.court_ruling_text_1 || 'No text'}`
        )
        .join('\n\n') || 'هیچ حکمی یافت نشد';

    formsText =
      formDocs
        .slice(0, 5)
        .map(
          (form) =>
            `Title: ${form.title || 'No title'}\nContent: ${form.content || 'No content'}`
        )
        .join('\n\n') || 'هیچ فرمی یافت نشد';

    terminologyText =
      matchedTerms
        .map((term) => `Term: ${term.term}\nDefinition: ${term.definition}`)
        .join('\n\n') || 'هیچ اصطلاحی یافت نشد';

    const customFormText = customForm
      ? `عنوان: ${customForm.title}\nمحتوا: ${customForm.content}`
      : 'هیچ فرم سفارشی یافت نشد';

    console.log('lawsText sample:', lawsText.substring(0, 100));
    console.log('courtRulingsText sample:', courtRulingsText.substring(0, 100));
    console.log('formsText sample:', formsText.substring(0, 100));
    console.log('terminologyText sample:', terminologyText.substring(0, 100));
    console.log('customFormText sample:', customFormText.substring(0, 100));

    const step4Prompt = `
شما یک وکیل حرفه‌ای پارسی هستید و کاربر مشتری شما است. هدف اصلی شما حفاظت از منافع کاربر از دیدگاه قانونی ایران است، با فرض اینکه کاربر بی‌گناه است یا تحت فشار و استرس شدید قرار دارد و نمی‌تواند واقعیت را به طور کامل بیان کند. پاسخ‌های عمومی و کلی غیرقابل‌قبول هستند. شما باید پاسخ خود را بر اساس داده‌های ارائه‌شده و جستجوی اینترنتی (در صورت نیاز) به‌صورت دقیق و با ارجاع به مواد قانونی، احکام قضایی، و فرم‌های حقوقی مرتبط تنظیم کنید.

داده‌های موجود:
- قوانین (تا ${maxLaws} مورد):
${lawsText}

- احکام قضایی (تا 20 مورد):
${courtRulingsText}

- فرم‌های حقوقی (تا 5 مورد):
${formsText}

- فرم سفارشی (در صورت وجود):
${customFormText}

- اصطلاحات قانونی (تا ۱۰۰ مورد):
${terminologyText}

- تاریخچه گفت‌وگو (۴ پیام آخر):
${conversation}

- پرس‌وجوی کاربر:
${userQuery}

**دستورالعمل‌های اجباری**:
1. **استفاده از داده‌های ارائه‌شده**: شما باید به‌طور خاص به قوانین، احکام قضایی، و اصطلاحات ارائه‌شده ارجاع دهید (مثلاً ماده ۱۱۳۰ قانون مدنی یا رای شماره ۹۸۸ دیوان عالی). اگر داده‌ها کافی نباشند، از ابزار جستجو برای یافتن قوانین یا احکام مرتبط با حقوق ایران استفاده کنید و منبع دقیق (مانند شماره ماده یا URL وب‌سایت معتبر) را ذکر کنید.
2. **فرم‌های حقوقی**: اگر کاربر درخواست تنظیم فرمی کرده است، فرم را به‌طور کامل بر اساس داده‌های موجود یا قالب‌های استاندارد ایرانی تنظیم کنید. اگر فرم سفارشی وجود دارد، راهنمایی کنید که چگونه از آن استفاده کند یا آن را پر کند. ما این فرم را به صورت فایل ورد به کاربر ارسال خواهیم کرد.
3. **جستجوی اینترنتی**: اگر اطلاعات کافی نیست (مثلاً فرم یا قانون خاص یافت نشد)، از ابزار جستجو برای یافتن اطلاعات به‌روز استفاده کنید (مانند وب‌سایت‌های رسمی قوه قضاییه: www.adliran.ir یا سامانه ملی قوانین: dotic.ir).
4. **پاسخ غیرکلی**: پاسخ‌های کلی مانند "با وکیل مشورت کنید" بدون جزئیات غیرقابل‌قبول هستند. هر بخش باید دقیق، با ارجاعات مشخص، و مرتبط با پرس‌وجوی کاربر باشد.

**ساختار پاسخ (به زبان پارسی)**:
1. **تعریف مشکل**: مشکل قانونی کاربر را بر اساس پرس‌وجو و زمینه آن توضیح دهید.
2. **نیازهای کاربر**: مشخص کنید کاربر به چه نوع کمکی (مانند تنظیم دادخواست، مشاوره حقوقی، جمع‌آوری مدارک) نیاز دارد.
3. **تحلیل حقوقی**: تحلیل جامعی ارائه دهید:
   - ارجاع مستقیم به قوانین (مانند ماده قانونی) و احکام قضایی (مانند شماره رای).
   - اگر کاربر اطلاعاتی داده که او را مجرم نشان می‌دهد، موارد را تک‌به‌تک بررسی کنید.
   - اگر حقوق کاربر تضییع شده، حقوق او را به‌صورت دقیق مشخص کنید.
   - از جستجوی اینترنتی برای قوانین یا احکام جدید استفاده کنید و ارجاع دهید.
4. **مراحل اقدام**: دستورالعمل گام‌به‌گام برای کاربر ارائه دهید (مانند ثبت دادخواست در سامانه ثنا).
5. **سوالات پیگیری**: سؤالاتی برای روشن شدن موقعیت بپرسید:
   - اگر کاربر متهم به جرمی است، صحت مدارک و شواهد را زیر سؤال ببرید (مانند شرایط شاهد معتبر).
   - اگر حقوق کاربر تضییع شده، سؤالاتی درباره جمع‌آوری مدارک بپرسید.
6. **شواهد**: راهنمایی کنید چه مدارکی (مانند سند، شاهد، گزارش پزشکی) معتبر هستند و چگونه جمع‌آوری شوند.
7. **فرم‌های حقوقی**: اگر کاربر فرم خاصی درخواست کرده، متن کامل فرم را تنظیم کنید یا راهنمایی گام‌به‌گام ارائه دهید. اگر فرم سفارشی وجود دارد، توضیح دهید که چگونه از آن استفاده کند.
8. **پیشنهاد وکیل**: نوع وکیل (مانند وکیل خانواده) را پیشنهاد دهید و بگویید کاربر چه اطلاعاتی به او بدهد.
9. **سامانه‌های اینترنتی**: سامانه‌های مرتبط (مانند سامانه ثنا: sana.adliran.ir) را معرفی کنید و نحوه استفاده را توضیح دهید.

پاسخ را به زبان پارسی، با لحن همدلانه و اطمینان‌بخش ارائه دهید و منافع کاربر را در اولویت قرار دهید.
`;

    console.log('Step 4 prompt length:', step4Prompt.length);
    console.log('Step 4 prompt sample:', step4Prompt.substring(0, 500));
    console.log(
      'Step 4 prompt includes lawsText?',
      step4Prompt.includes(lawsText.substring(0, 50))
    );
    console.log(
      'Step 4 prompt includes courtRulingsText?',
      step4Prompt.includes(courtRulingsText.substring(0, 50))
    );

    let step4Response = 'متأسفیم، در تولید پاسخ حقوقی مشکلی پیش آمد.';
    try {
      console.log('Sending final advice prompt to AI...');
      const step4Result = await getGenerativeModel(
        step4Prompt,
        'Step 4: Final Advice',
        'gemini-2.0-flash',
        4
      );
      step4Response = step4Result.text;
      console.log('Generated legal advice');
      console.log('Response length:', step4Response.length);
    } catch (e) {
      console.error('Step 4 failed:', e.message);
    }

    console.log('----------------------------------------');
    console.log('PREPARING EXCEL OUTPUT');
    console.log('----------------------------------------');
    console.log(
      'Court ruling documents for Excel:',
      courtRulingDocuments
        .map((doc) => ({
          title: doc.title,
          ruling_group: doc.ruling_group,
          text: doc.court_ruling_text_1?.substring(0, 100) || 'No text',
        }))
        .slice(0, 3)
    );
    console.log(
      'Form documents for Excel:',
      formDocs
        .map((doc) => ({
          title: doc.title,
          content: doc.content?.substring(0, 100) || 'No content',
        }))
        .slice(0, 3)
    );

    return {
      is_meaningful: true,
      step4Response,
      laws: finalLaws,
      courtRulingDocuments,
      matchedTerms,
      formDocs,
      customForm,
    };
  } catch (e) {
    console.error('==================================');
    console.error('UNEXPECTED ERROR:', e.message);
    console.error('==================================');
    return {
      is_meaningful: false,
      error: 'Internal processing error',
      step4Response: 'متأسفیم، خطایی در پردازش رخ داد.',
      laws: [],
      courtRulingDocuments: [],
      matchedTerms: [],
      formDocs: [],
      customForm: null,
    };
  }
};

// Main function
export default async ({ req, res, log, error }) => {
  log('==================================');
  log('LEGAL QUERY FUNCTION STARTED');
  log('==================================');

  // Initialize Appwrite client
  const client = new Client()
    .setEndpoint(
      process.env.APPWRITE_ENDPOINT || 'https://cloud.appwrite.io/v1'
    )
    .setProject(process.env.APPWRITE_PROJECT_ID)
    .setKey(process.env.APPWRITE_API_KEY);

  const databases = new Databases(client);
  const storage = new Storage(client);

  try {
    // Validate environment variables
    const requiredEnvVars = [
      'APPWRITE_PROJECT_ID',
      'APPWRITE_API_KEY',
      'APPWRITE_DATABASE_ID',
      'APPWRITE_ENDPOINT',
      'APPWRITE_BUCKET_ID',
      'GEMINI_API_KEY',
    ];

    const missingVars = requiredEnvVars.filter(
      (varName) => !process.env[varName]
    );
    if (missingVars.length > 0) {
      error(`Missing environment variables: ${missingVars.join(', ')}`);
      return res.json(
        { error: `Missing environment variables: ${missingVars.join(', ')}` },
        { status: 500 }
      );
    }

    // Parse request body
    let requestBody;
    try {
      requestBody =
        typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
      log('Received request body');
    } catch (e) {
      error('Failed to parse request body:', e.message);
      return res.json({ error: 'Invalid request body' }, { status: 400 });
    }

    const { clerkId, text, sessionId } = requestBody;
    if (!clerkId || !text || !sessionId) {
      error('Missing required fields: clerkId, text, or sessionId');
      return res.json(
        { error: 'Missing required fields: clerkId, text, or sessionId' },
        { status: 400 }
      );
    }

    log(`Processing query for clerkId: ${clerkId}, sessionId: ${sessionId}`);
    log(`Query text: ${text.substring(0, 50)}...`);

    // User management
    let user;
    let activeSession;
    try {
      log('Checking user existence...');
      const userResponse = await databases.listDocuments(
        process.env.APPWRITE_DATABASE_ID,
        process.env.APPWRITE_USERS_COLLECTION_ID || 'users',
        [Query.equal('clerkId', clerkId)]
      );

      if (userResponse.documents.length === 0) {
        log('Creating new user...');
        user = await databases.createDocument(
          process.env.APPWRITE_DATABASE_ID,
          process.env.APPWRITE_USERS_COLLECTION_ID || 'users',
          ID.unique(),
          {
            clerkId,
            telegram_id: null,
            is_answer_in_progress: false,
            usage_this_month: 0,
            usage_total: 0,
            is_blocked: false,
          }
        );
        log(`Created user ${user.$id}`);
      } else {
        user = userResponse.documents[0];
        log(`Found user ${user.$id}`);
      }

      if (user.is_blocked) {
        log(`User ${clerkId} is blocked`);
        return res.json({ error: 'کاربر مسدود شده است' }, { status: 403 });
      }

      if (user.usage_this_month >= 4) {
        log('User exceeded monthly limit');
        return res.json(
          {
            error:
              'شما به سقف استفاده هفتگی (۴ درخواست) رسیده‌اید. لطفاً هفته بعد دوباره تلاش کنید.',
          },
          { status: 429 }
        );
      }

      const lastUpdateTime = new Date(user.$updatedAt).getTime();
      const currentTime = Date.now();
      const timeSinceLastUpdate = (currentTime - lastUpdateTime) / 1000;
      log(`Time since last update: ${timeSinceLastUpdate} seconds`);

      if (timeSinceLastUpdate < 120) {
        log('User request too frequent');
        return res.json(
          { error: 'لطفاً ۱۲۰ ثانیه بین درخواست‌ها فاصله دهید.' },
          { status: 429 }
        );
      }

      // Session management
      log('Checking session existence...');
      const sessionResponse = await databases.listDocuments(
        process.env.APPWRITE_DATABASE_ID,
        process.env.APPWRITE_SESSIONS_COLLECTION_ID || 'chat_sessions',
        [Query.equal('$id', sessionId), Query.equal('user_ref', user.$id)]
      );

      if (sessionResponse.documents.length === 0) {
        log('Invalid sessionId, creating new session');
        // Deactivate existing active sessions
        const activeSessions = await databases.listDocuments(
          process.env.APPWRITE_DATABASE_ID,
          process.env.APPWRITE_SESSIONS_COLLECTION_ID || 'chat_sessions',
          [Query.equal('user_ref', user.$id), Query.equal('is_active', true)]
        );
        for (const session of activeSessions.documents) {
          await databases.updateDocument(
            process.env.APPWRITE_DATABASE_ID,
            process.env.APPWRITE_SESSIONS_COLLECTION_ID || 'chat_sessions',
            session.$id,
            { is_active: false }
          );
          log(`Deactivated session ${session.$id}`);
        }

        activeSession = await databases.createDocument(
          process.env.APPWRITE_DATABASE_ID,
          process.env.APPWRITE_SESSIONS_COLLECTION_ID || 'chat_sessions',
          ID.unique(),
          {
            user_ref: user.$id,
            is_active: true,
          }
        );
        log(`Created new session ${activeSession.$id}`);
      } else {
        activeSession = sessionResponse.documents[0];
        log(`Using existing session ${activeSession.$id}`);
      }

      log('Saving user message...');
      await databases.createDocument(
        process.env.APPWRITE_DATABASE_ID,
        process.env.APPWRITE_CHATS_COLLECTION_ID || 'chats',
        ID.unique(),
        {
          chat_session_ref: activeSession.$id,
          message: text,
          role: 'user',
        }
      );
      log('User message saved');
    } catch (e) {
      error('User/session management failed:', e.message);
      return res.json(
        { error: `خطا در مدیریت کاربر/جلسه: ${e.message}` },
        { status: 500 }
      );
    }

    // Retrieve conversation history
    let conversation = '';
    try {
      log('Retrieving conversation history...');
      const chatMessages = await databases.listDocuments(
        process.env.APPWRITE_DATABASE_ID,
        process.env.APPWRITE_CHATS_COLLECTION_ID || 'chats',
        [
          Query.equal('chat_session_ref', activeSession.$id),
          Query.orderDesc('$createdAt'),
          Query.limit(4),
        ]
      );

      conversation = chatMessages.documents
        .map((msg) => `${msg.role}: ${msg.message}`)
        .reverse()
        .join('\n');

      log(
        `Conversation history (${chatMessages.documents.length} messages):\n${conversation.substring(0, 100)}...`
      );
    } catch (e) {
      error('Failed to retrieve conversation history:', e.message);
    }

    // Process legal query
    log('Processing legal query...');
    const lawyerResponse = await processLegalQuery(
      text,
      conversation,
      databases
    );

    const {
      is_meaningful,
      step4Response,
      laws = [],
      courtRulingDocuments = [],
      matchedTerms = [],
      formDocs = [],
      customForm,
    } = lawyerResponse;

    if (!is_meaningful) {
      log('Conversation not meaningful');
      return res.json(
        { error: 'لطفاً سؤال حقوقی واضحی مطرح کنید.' },
        { status: 400 }
      );
    }

    // Update usage counters
    try {
      log('Updating user usage counters...');
      await databases.updateDocument(
        process.env.APPWRITE_DATABASE_ID,
        process.env.APPWRITE_USERS_COLLECTION_ID || 'users',
        user.$id,
        {
          usage_this_month: user.usage_this_month + 1,
          usage_total: user.usage_total + 1,
        }
      );
      log('Usage counters updated');
    } catch (e) {
      error('Failed to update usage counters:', e.message);
    }

    // Prepare response
    let wordDocument = null;
    let excelFile = null;
    let forms = [];

    // Generate Word document
    if (step4Response) {
      try {
        log('Generating Word document...');
        const wordBuffer = await createWordDocument(step4Response);
        wordDocument = wordBuffer.toString('base64');
        log(`Word document generated, size: ${wordBuffer.length} bytes`);
      } catch (e) {
        error('Word document generation failed:', e.message);
      }
    }

    // Generate Excel file
    try {
      log('Generating Excel report...');
      const workbook = XLSX.utils.book_new();

      if (laws.length > 0) {
        const lawsSheet = XLSX.utils.json_to_sheet(
          laws.map((law) => ({
            ID: law.$id || '',
            Title: law.title || 'No title',
            Subcategory: law.subcategory || 'No subcategory',
            Content: law.content || 'No content',
          }))
        );
        XLSX.utils.book_append_sheet(workbook, lawsSheet, 'قوانین');
        log('Added laws sheet');
      }

      if (courtRulingDocuments.length > 0) {
        const courtRulingsSheet = XLSX.utils.json_to_sheet(
          courtRulingDocuments.map((ruling) => ({
            ID: ruling.$id || '',
            Title: ruling.title || 'No title',
            RulingGroup: ruling.ruling_group || 'No group',
            Title1: ruling.court_ruling_title_1 || 'No title',
            Text1: ruling.court_ruling_text_1 || 'No text',
            Title2: ruling.court_ruling_title_2 || 'No title',
            Text2: ruling.court_ruling_text_2 || 'No text',
            SourceURL: ruling.source_url || 'No URL',
          }))
        );
        XLSX.utils.book_append_sheet(workbook, courtRulingsSheet, 'آرا دادگاه');
        log('Added rulings sheet');
      }

      if (matchedTerms.length > 0) {
        const terminologySheet = XLSX.utils.json_to_sheet(
          matchedTerms.map((term) => ({
            Term: term.term || 'No term',
            Definition: term.definition || 'No definition',
          }))
        );
        XLSX.utils.book_append_sheet(
          workbook,
          terminologySheet,
          'اصطلاحات حقوقی'
        );
        log('Added terms sheet');
      }

      if (formDocs.length > 0) {
        const formsSheet = XLSX.utils.json_to_sheet(
          formDocs.map((form) => ({
            Title: form.title || 'No title',
            Content: form.content || 'No content',
          }))
        );
        XLSX.utils.book_append_sheet(workbook, formsSheet, 'فرم‌ها');
        log('Added forms sheet');
      }

      const excelBuffer = XLSX.write(workbook, {
        bookType: 'xlsx',
        type: 'buffer',
      });
      excelFile = excelBuffer.toString('base64');
      log('Excel file generated');
    } catch (e) {
      error('Excel file generation failed:', e.message);
    }

    // Fetch and prepare forms
    let formsSent = 0;
    log(`Processing ${Math.min(formDocs.length, 2)} form files...`);
    let formsToSend = formDocs.slice(0, 2);
    if (customForm) {
      try {
        log(`Generating custom form: ${customForm.title}`);
        const customFormBuffer = await createWordDocument(customForm.content);
        forms.push({
          title: customForm.title,
          fileData: customFormBuffer.toString('base64'),
        });
        log(`Custom form added: ${customForm.title}`);
        formsSent++;
      } catch (e) {
        error(`Custom form generation failed: ${e.message}`);
      }
    }

    for (const form of formsToSend) {
      try {
        if (!form.url) {
          log(`Form ${form.title} missing URL`);
          continue;
        }

        const fileIdMatch = form.url.match(/files\/([^\/]+)/);
        if (!fileIdMatch?.[1]) {
          log(`Invalid URL format for ${form.title}: ${form.url}`);
          continue;
        }

        const fileId = fileIdMatch[1];
        const bucketId = process.env.APPWRITE_BUCKET_ID || 'default';

        const { fileData } = await getAppwriteFile(
          bucketId,
          fileId,
          form.title,
          form.url
        );

        forms.push({
          title: form.title,
          fileData: fileData.toString('base64'),
        });
        log(`Form ${form.title} added`);
        formsSent++;
      } catch (e) {
        error(`Form ${form.title} failed: ${e.message}`);
      }
    }
    log(`Processed ${formsSent} form files`);

    // Save AI response
    if (step4Response) {
      try {
        log('Saving AI response...');
        await databases.createDocument(
          process.env.APPWRITE_DATABASE_ID,
          process.env.APPWRITE_CHATS_COLLECTION_ID || 'chats',
          ID.unique(),
          {
            chat_session_ref: activeSession.$id,
            message: step4Response,
            role: 'assistant',
          }
        );
        log('AI response saved');
      } catch (e) {
        error('Failed to save AI response:', e.message);
      }
    }

    // Return response
    log('Returning response to client');
    return res.json(
      {
        message: step4Response,
        wordDocument: wordDocument || null,
        excelFile: excelFile || null,
        forms,
      },
      { status: 200 }
    );
  } catch (e) {
    error('==================================');
    error('GLOBAL ERROR:', e.message);
    error('==================================');
    return res.json(
      { error: `خطای غیرمنتظره رخ داد: ${e.message}` },
      { status: 500 }
    );
  }
};

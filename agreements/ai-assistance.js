/* Ovateq Docs Agreement AI — Phase 2 provider-neutral integration boundary.
   No AI provider, API key, network endpoint or generated response is included. */
(function (global) {
  'use strict';

  const SCHEMA_VERSION = 'ovateq.agreement-ai.v1';
  const VERSION = 'phase2-boundary-20260927';
  const CAPABILITIES = Object.freeze([
    'identify_agreement_type',
    'extract_structured_fields',
    'find_missing_information',
    'review_agreement',
    'suggest_wording',
    'explain_agreement'
  ]);
  const SOURCES = Object.freeze(['extracted_from_user_input', 'ai_suggestion', 'not_provided']);
  const BLOCKED_PATHS = /^(id|number|status|templateVersion|createdAt|updatedAt|history|attachments|signatures)(\.|$)/;
  const BASE_PATHS = [
    'title', 'date', 'currency', 'governingLaw', 'amount', 'deposit', 'interestRate',
    'paymentMethod', 'installments', 'frequency', 'firstPaymentDate', 'specialConditions',
    'partyA.name', 'partyA.idType', 'partyA.idNumber', 'partyA.phone', 'partyA.email', 'partyA.address',
    'partyB.name', 'partyB.idType', 'partyB.idNumber', 'partyB.phone', 'partyB.email', 'partyB.address',
    'witnesses.0.name', 'witnesses.0.phone', 'witnesses.0.idNumber',
    'witnesses.1.name', 'witnesses.1.phone', 'witnesses.1.idNumber'
  ];

  let provider = null;

  const isObject = value => !!value && typeof value === 'object' && !Array.isArray(value);
  const text = value => typeof value === 'string' ? value.trim() : '';
  const cleanText = (value, max = 12000) => text(value).slice(0, max);
  const id = () => global.crypto?.randomUUID?.() || `ai-${Date.now()}-${Math.random().toString(36).slice(2)}`;

  function providerStatus() {
    if (!provider) {
      return Object.freeze({
        available: false,
        status: 'unavailable',
        code: 'AI_PROVIDER_NOT_CONFIGURED',
        providerId: null,
        privacyNotice: '',
        message: 'AI assistance is not connected yet. You can continue creating this agreement manually.'
      });
    }
    return Object.freeze({
      available: true,
      status: 'available',
      code: 'AI_PROVIDER_CONFIGURED',
      providerId: provider.id,
      privacyNotice: provider.privacyNotice,
      message: 'AI assistance is available. Review every result before applying it.'
    });
  }

  function registerProvider(candidate) {
    if (!isObject(candidate)) throw new TypeError('Agreement AI provider must be an object.');
    if (!/^[a-z0-9][a-z0-9._-]{1,63}$/i.test(text(candidate.id))) throw new TypeError('Agreement AI provider requires a stable id.');
    if (typeof candidate.execute !== 'function') throw new TypeError('Agreement AI provider requires execute(request).');
    if (!text(candidate.privacyNotice)) throw new TypeError('Agreement AI provider requires an honest privacy notice.');
    const declared = Array.isArray(candidate.capabilities) ? candidate.capabilities : CAPABILITIES;
    const unsupported = declared.filter(item => !CAPABILITIES.includes(item));
    if (unsupported.length) throw new TypeError(`Unsupported Agreement AI capabilities: ${unsupported.join(', ')}`);
    provider = Object.freeze({
      id: text(candidate.id),
      privacyNotice: cleanText(candidate.privacyNotice, 1000),
      capabilities: Object.freeze([...new Set(declared)]),
      execute: candidate.execute
    });
    return providerStatus();
  }

  function unregisterProvider() {
    provider = null;
    return providerStatus();
  }

  function templateIds(catalogue) {
    return new Set((Array.isArray(catalogue) ? catalogue : []).map(item => text(item?.id)).filter(Boolean));
  }

  function allowedPaths(catalogue, templateId) {
    const paths = new Set(BASE_PATHS);
    const template = (Array.isArray(catalogue) ? catalogue : []).find(item => item?.id === templateId);
    for (const field of template?.fields || []) if (text(field?.name)) paths.add(`details.${text(field.name)}`);
    return paths;
  }

  function responseContract(capability, catalogue) {
    const ids = [...templateIds(catalogue)];
    const common = {
      schemaVersion: SCHEMA_VERSION,
      requestId: 'Copy the requestId exactly.',
      capability,
      providerId: 'Configured provider identifier.',
      generatedAt: 'ISO-8601 timestamp.',
      disclaimer: 'AI output is a suggestion, not legal advice.'
    };
    const contracts = {
      identify_agreement_type: {
        ...common,
        result: { templateId: `One of: ${ids.join(', ')}`, confidence: 'Number from 0 to 1', rationale: 'Short explanation', evidence: ['Exact phrases from user input'] }
      },
      extract_structured_fields: {
        ...common,
        result: {
          templateId: `One of: ${ids.join(', ')}`,
          fields: [{ path: 'Existing Agreement field path', value: 'String or number', source: 'extracted_from_user_input', evidence: 'Exact supporting user text' }],
          missing: [{ path: 'Existing Agreement field path', label: 'Human label', reason: 'Why it is needed', source: 'not_provided' }]
        }
      },
      find_missing_information: {
        ...common,
        result: { missing: [{ path: 'Existing Agreement field path', label: 'Human label', reason: 'Why it is needed', source: 'not_provided' }] }
      },
      review_agreement: {
        ...common,
        result: { findings: [{ id: 'Stable finding id', level: 'information | potential_issue | incomplete', title: 'Short title', message: 'Careful non-legal conclusion', paths: ['Related field paths'] }] }
      },
      suggest_wording: {
        ...common,
        result: { suggestions: [{ id: 'Stable suggestion id', targetPath: 'Existing editable text field', existing: 'Current text', suggested: 'Proposed text', reason: 'Why it may be clearer', source: 'ai_suggestion' }] }
      },
      explain_agreement: {
        ...common,
        result: { summary: 'Plain-language explanation', sections: [{ sectionId: 'Stable section id', title: 'Section title', explanation: 'What the document says without legal guarantees' }] }
      }
    };
    return contracts[capability];
  }

  function sanitiseAgreement(agreement) {
    if (!isObject(agreement)) return null;
    const copy = typeof structuredClone === 'function' ? structuredClone(agreement) : JSON.parse(JSON.stringify(agreement));
    delete copy.attachments;
    delete copy.history;
    delete copy.signatures;
    delete copy.createdAt;
    delete copy.updatedAt;
    return copy;
  }

  function createRequest({ capability, userInput = '', agreement = null, catalogue = [], locale = 'en-GH' } = {}) {
    if (!CAPABILITIES.includes(capability)) throw new TypeError('Unknown Agreement AI capability.');
    const templates = (Array.isArray(catalogue) ? catalogue : []).map(template => ({
      id: text(template?.id),
      family: cleanText(template?.family, 120),
      title: cleanText(template?.title, 160),
      roles: Array.isArray(template?.roles) ? template.roles.map(role => cleanText(role, 120)) : [],
      fields: Array.isArray(template?.fields) ? template.fields.map(field => ({
        name: text(field?.name), label: cleanText(field?.label, 160), type: text(field?.type) || 'text',
        required: !!field?.required, options: Array.isArray(field?.options) ? field.options.map(option => cleanText(option, 120)) : []
      })) : []
    })).filter(template => template.id);
    return Object.freeze({
      schemaVersion: SCHEMA_VERSION,
      requestId: id(),
      capability,
      locale: text(locale) || 'en-GH',
      createdAt: new Date().toISOString(),
      userInput: cleanText(userInput),
      agreement: sanitiseAgreement(agreement),
      templateCatalogue: templates,
      instructions: {
        neverInventFacts: true,
        returnStructuredDataOnly: true,
        userApprovalRequiredBeforeChanges: true,
        preserveExistingAgreementEngine: true,
        legalAdviceProhibited: true,
        allowedSources: SOURCES
      },
      responseContract: responseContract(capability, templates)
    });
  }

  function validatePath(path, paths) {
    return !!text(path) && !BLOCKED_PATHS.test(path) && paths.has(path);
  }

  function validateResponse(request, raw) {
    const errors = [];
    if (!isObject(raw)) return { valid: false, errors: ['Provider response must be an object.'], value: null };
    if (raw.schemaVersion !== SCHEMA_VERSION) errors.push('Response schemaVersion does not match the request.');
    if (raw.requestId !== request.requestId) errors.push('Response requestId does not match the request.');
    if (raw.capability !== request.capability) errors.push('Response capability does not match the request.');
    if (!isObject(raw.result)) errors.push('Response result must be an object.');
    const result = isObject(raw.result) ? raw.result : {};
    const ids = templateIds(request.templateCatalogue);
    const chosenTemplate = text(result.templateId) || text(request.agreement?.type);
    if (result.templateId && !ids.has(text(result.templateId))) errors.push('Response contains an unknown Agreement template ID.');
    const paths = allowedPaths(request.templateCatalogue, chosenTemplate);

    if (request.capability === 'identify_agreement_type') {
      if (!ids.has(text(result.templateId))) errors.push('A valid templateId is required.');
      if (!(Number(result.confidence) >= 0 && Number(result.confidence) <= 1)) errors.push('confidence must be between 0 and 1.');
      if (!text(result.rationale)) errors.push('A rationale is required.');
    }

    if (request.capability === 'extract_structured_fields') {
      if (!ids.has(text(result.templateId))) errors.push('A valid templateId is required.');
      if (!Array.isArray(result.fields)) errors.push('fields must be an array.');
      for (const field of result.fields || []) {
        if (!validatePath(field?.path, paths)) errors.push(`Unsafe or unknown extracted field path: ${text(field?.path) || '(empty)'}.`);
        if (!['string', 'number'].includes(typeof field?.value)) errors.push(`Extracted value for ${text(field?.path)} must be text or a number.`);
        if (field?.source !== 'extracted_from_user_input') errors.push(`Extracted field ${text(field?.path)} must use source extracted_from_user_input.`);
        if (!text(field?.evidence)) errors.push(`Extracted field ${text(field?.path)} requires evidence from the user's input.`);
      }
    }

    if (['extract_structured_fields', 'find_missing_information'].includes(request.capability)) {
      if (!Array.isArray(result.missing)) errors.push('missing must be an array.');
      for (const item of result.missing || []) {
        if (!validatePath(item?.path, paths)) errors.push(`Unsafe or unknown missing-information path: ${text(item?.path) || '(empty)'}.`);
        if (item?.source !== 'not_provided') errors.push(`Missing item ${text(item?.path)} must use source not_provided.`);
        if (!text(item?.label) || !text(item?.reason)) errors.push(`Missing item ${text(item?.path)} requires a label and reason.`);
      }
    }

    if (request.capability === 'review_agreement') {
      if (!Array.isArray(result.findings)) errors.push('findings must be an array.');
      for (const finding of result.findings || []) {
        if (!['information', 'potential_issue', 'incomplete'].includes(finding?.level)) errors.push('Review finding has an unsupported level.');
        if (!text(finding?.title) || !text(finding?.message)) errors.push('Review finding requires a title and message.');
        for (const path of finding?.paths || []) if (!validatePath(path, paths)) errors.push(`Review finding contains unknown field path: ${text(path)}.`);
      }
    }

    if (request.capability === 'suggest_wording') {
      if (!Array.isArray(result.suggestions)) errors.push('suggestions must be an array.');
      for (const suggestion of result.suggestions || []) {
        if (!validatePath(suggestion?.targetPath, paths)) errors.push(`Unsafe or unknown suggestion target: ${text(suggestion?.targetPath) || '(empty)'}.`);
        if (!text(suggestion?.suggested) || !text(suggestion?.reason)) errors.push('Wording suggestion requires proposed text and a reason.');
        if (suggestion?.source !== 'ai_suggestion') errors.push('Wording suggestion must use source ai_suggestion.');
      }
    }

    if (request.capability === 'explain_agreement') {
      if (!text(result.summary)) errors.push('Explanation summary is required.');
      if (!Array.isArray(result.sections)) errors.push('Explanation sections must be an array.');
      for (const section of result.sections || []) if (!text(section?.title) || !text(section?.explanation)) errors.push('Each explanation section requires a title and explanation.');
    }

    return { valid: errors.length === 0, errors: [...new Set(errors)], value: errors.length ? null : raw };
  }

  async function run(options) {
    const status = providerStatus();
    if (!status.available) return { ok: false, ...status };
    const request = createRequest(options);
    if (!provider.capabilities.includes(request.capability)) {
      return { ok: false, status: 'unavailable', code: 'AI_CAPABILITY_NOT_SUPPORTED', providerId: provider.id, message: 'The configured AI provider does not support this action.' };
    }
    try {
      const raw = await provider.execute(request);
      const checked = validateResponse(request, raw);
      if (!checked.valid) return { ok: false, status: 'invalid_response', code: 'AI_RESPONSE_REJECTED', providerId: provider.id, request, errors: checked.errors, message: 'AI returned data that did not match the safe Agreement schema. Nothing was changed.' };
      return { ok: true, status: 'complete', code: 'AI_RESPONSE_ACCEPTED', providerId: provider.id, privacyNotice: provider.privacyNotice, request, response: checked.value };
    } catch (error) {
      return { ok: false, status: 'failed', code: 'AI_PROVIDER_FAILED', providerId: provider.id, request, message: 'AI assistance is temporarily unavailable. You can continue creating this agreement manually.', errorName: text(error?.name) || 'Error' };
    }
  }

  const api = Object.freeze({
    version: VERSION,
    schemaVersion: SCHEMA_VERSION,
    capabilities: CAPABILITIES,
    sources: SOURCES,
    providerStatus,
    registerProvider,
    unregisterProvider,
    createRequest,
    validateResponse,
    run
  });

  global.OvateqAgreementAI = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);

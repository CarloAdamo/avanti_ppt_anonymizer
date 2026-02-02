// Avanti PPT Anonymizer — Entry point

import { extractAllText } from './extraction.js';
import { classifyLocally, classifyWithAI, fallbackClassify, buildRewrites } from './classification.js';
import { replaceAllShapes } from './replacement.js';
import { showDoneSection, showLoading, hideLoading, showStatus, hideStatus } from './ui.js';

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log('Avanti Anonymizer loaded');
        initializeUI();
    }
});

function initializeUI() {
    document.getElementById('scan-btn').addEventListener('click', anonymizePresentation);
    document.getElementById('new-scan-btn').addEventListener('click', resetAndAnonymize);
}

async function anonymizePresentation() {
    hideStatus();

    try {
        showLoading('Extraherar text...');
        const slideData = await extractAllText();

        if (slideData.length === 0) {
            hideLoading();
            showStatus('Ingen text hittades i presentationen.', 'info');
            return;
        }

        showLoading('Klassificerar text...');
        const { classified, unclassified } = classifyLocally(slideData);

        let aiClassifications = [];
        if (unclassified.length > 0) {
            try {
                showLoading('AI klassificerar text...');
                aiClassifications = await classifyWithAI(unclassified);
            } catch (error) {
                console.warn('AI classification unavailable, using fallback:', error);
                aiClassifications = fallbackClassify(unclassified);
            }
        }

        const allClassifications = [...classified, ...aiClassifications];

        showLoading('Ersätter text...');
        const rewrites = buildRewrites(slideData, allClassifications);

        if (rewrites.length === 0) {
            hideLoading();
            showStatus('Ingen text behövde generaliseras.', 'success');
            return;
        }

        const count = await replaceAllShapes(slideData, rewrites);

        hideLoading();
        showDoneSection(count);

    } catch (error) {
        hideLoading();
        console.error('Anonymization error:', error);
        showStatus('Ett fel uppstod: ' + error.message, 'error');
    }
}

function resetAndAnonymize() {
    document.getElementById('scan-section').classList.remove('hidden');
    document.getElementById('done-section').classList.add('hidden');
    hideStatus();
    anonymizePresentation();
}

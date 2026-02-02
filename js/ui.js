// Avanti PPT Anonymizer — UI helper functions

export function showDoneSection(count) {
    document.getElementById('scan-section').classList.add('hidden');
    document.getElementById('done-section').classList.remove('hidden');
    document.getElementById('replaced-count').textContent = count;
}

export function showLoading(text) {
    document.getElementById('loading-text').textContent = text;
    document.getElementById('loading').classList.remove('hidden');
}

export function hideLoading() {
    document.getElementById('loading').classList.add('hidden');
}

export function showStatus(message, type) {
    const status = document.getElementById('status');
    status.textContent = message;
    status.className = `status ${type}`;
    status.classList.remove('hidden');
}

export function hideStatus() {
    document.getElementById('status').classList.add('hidden');
}

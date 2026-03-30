const fs = require('fs');
const path = require('path');

// original_sermon 폴더에서 모든 파일 읽기
const sourceDir = path.join(__dirname, '..', 'original_sermon');
const outputDir = path.join(__dirname, '..', 'sermons_for_lm');

// 출력 폴더 생성
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

// 모든 파일 목록 가져오기
const files = fs.readdirSync(sourceDir)
    .filter(f => f.endsWith('.txt') || f.endsWith('.TXT'))
    .sort();

console.log(`전체 설교 파일 수: ${files.length}`);

// 300개 파일로 나누기
const numChunks = 300;
const chunkSize = Math.ceil(files.length / numChunks);
console.log(`각 청크당 약 ${chunkSize}개 파일`);

// 파일 내용 읽기 및 인코딩 확인
let allContent = [];
files.forEach((file, index) => {
    const filePath = path.join(sourceDir, file);
    let content = fs.readFileSync(filePath, 'utf8');
    // BOM 제거 (存在할 경우)
    if (content.charCodeAt(0) === 0xFEFF) {
        content = content.slice(1);
    }
    allContent.push({ file, content });
    if ((index + 1) % 100 === 0) {
        console.log(`${index + 1}개 파일 읽기 완료`);
    }
});

// 300개 파일로 나누기 (파일 단위)
const chunkedFiles = [];
for (let i = 0; i < allContent.length; i += chunkSize) {
    chunkedFiles.push(allContent.slice(i, i + chunkSize));
}

console.log(`총 ${chunkedFiles.length}개의 출력 파일 생성`);

// 각 청크를 파일로 저장
chunkedFiles.forEach((chunk, index) => {
    let combinedContent = '';
    
    chunk.forEach(item => {
        combinedContent += `=== ${item.file} ===\n\n`;
        combinedContent += item.content.trim();
        combinedContent += '\n\n';
        combinedContent += '='.repeat(50);
        combinedContent += '\n\n';
    });
    
    const outputFileName = `sermons_${String(index + 1).padStart(3, '0')}.txt`;
    const outputPath = path.join(outputDir, outputFileName);
    
    // UTF-8 with BOM으로 저장 (한글 호환성)
    const BOM = '\uFEFF';
    fs.writeFileSync(outputPath, BOM + combinedContent, 'utf8');
    
    console.log(`생성: ${outputFileName} (${chunk.length}개 설교 포함)`);
});

console.log(`\n완료! ${outputDir}에 ${chunkedFiles.length}개의 파일이 생성되었습니다.`);

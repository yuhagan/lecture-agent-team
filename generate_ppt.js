const pptxgen = require('pptxgenjs');
const fs = require('fs');

async function buildPpt() {
    let pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';

    // Slide 1
    let slide1 = pptx.addSlide();
    slide1.addText('상완골(Humerus)', { x: 0.5, y: 0.5, w: 4, h: 0.5, fontSize: 28, bold: true, color: '1E3A8A' });
    slide1.addText(
        '• 상완골 개요: 이것은 상완의 긴 뼈입니다.\n' +
        '• 근위단 지표 (Landmarks):\n' +
        '   - 상완골두: 견갑골의 관절와와 관절하는 부분\n' +
        '   - 대결절 (Greater tubercle): 외측 위치, 근육 부착\n' +
        '   - 소결절 (Lesser tubercle): 앞쪽 위치\n' +
        '   - 상완이두근구: 두 결절 사이를 지나가는 고랑\n' +
        '   - 해부경 (Anatomical neck): 상완골두 바로 밑 부위\n' +
        '• 골간 및 원위부:\n' +
        '   - 삼각근조면: 골간 중앙, 삼각근 위치\n' +
        '   - 골격 형태: 원위단으로 갈수록 삼각형 구조',
        { x: 0.5, y: 1.2, w: 4.5, h: 4, fontSize: 13, color: '333333' }
    );
    // Add cropped image right side (Approximated rect for diagram)
    // using sizing: { type: 'crop', ... }
    slide1.addImage({
        path: 'pages/page_1.png',
        x: 5.2, y: 0.5, w: 4.5, h: 6.5,
        sizing: { type: 'crop', x: 0, y: 1, w: 6, h: 6 } // approx right-center
    });

    // Slide 2
    let slide2 = pptx.addSlide();
    slide2.addText('관절상완관절 (Glenohumeral Joint)', { x: 0.5, y: 0.5, w: 5, h: 0.5, fontSize: 24, bold: true, color: '1E3A8A' });
    slide2.addText(
        '• 관절의 특징: 어깨의 주된 관절로 가동성이 매우 높지만 불안정한 구조\n' +
        '• 구조적 세부사항:\n' +
        '   - 상완골두: 공 모양 구면의 2/5, 내측 및 후상방 향함\n' +
        '   - 관절와 (Glenoid cavity): 상완골두보다 2~3배 작고 상당히 얕음\n' +
        '   - 방향: 주로 외측 및 약간 전상방 향함\n' +
        '• 안정화 요소:\n' +
        '   - 관절순 (Glenoid labrum): 나사받이처럼 관절을 봉합하여 안정성 높임',
        { x: 0.5, y: 1.2, w: 4.5, h: 4, fontSize: 13, color: '333333' }
    );
    slide2.addImage({
        path: 'pages/page_2.png',
        x: 5.2, y: 0.5, w: 4.5, h: 6.5,
        sizing: { type: 'crop', x: 2, y: 1, w: 5, h: 6 } // approx left-side pictures
    });

    // Slide 3
    let slide3 = pptx.addSlide();
    slide3.addText('관절낭 및 인대 (Capsule / Ligaments)', { x: 0.5, y: 0.5, w: 5.5, h: 0.5, fontSize: 22, bold: true, color: '1E3A8A' });
    slide3.addText(
        '• 관절낭 (Joint Capsule):\n' +
        '   - 견갑골 관절와 테두리 ~ 상완골 해부경\n' +
        '   - 매우 느슨하고 주름이 많으며 아래쪽이 가장 약함\n' +
        '• 주요 인대:\n' +
        '   - 오훼상완인대: 오훼돌기에서 대결절 주행\n' +
        '   - 관절상완인대: 관절와 테두리에서부터 주행\n' +
        '• 보강 및 탈구:\n' +
        '   - 전하방 부위가 가장 약해 탈구 흔함\n' +
        '• 휴식 위치 (Resting Position):\n' +
        '   - 상완이 가볍게 굴곡, 외전, 내회전된 상태 (최대 이완)',
        { x: 0.5, y: 1.2, w: 4.5, h: 5.5, fontSize: 12, color: '333333' }
    );
    slide3.addImage({
        path: 'pages/page_3.png',
        x: 5.2, y: 0.5, w: 4.5, h: 6.5,
        sizing: { type: 'crop', x: 1, y: 2, w: 6, h: 6 } 
    });

    await pptx.writeFile({ fileName: 'lecture-agent-team-presentation.pptx' });
    console.log("Created high quality PPTX.");
}

buildPpt();

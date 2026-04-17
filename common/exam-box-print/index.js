const { createApp, ref, computed, watch, onBeforeUnmount } = Vue;

const MAX_SUBJECTS = 20;
const PREVIEW_DEBOUNCE_MS = 300;
const DEFAULT_REGION = '嘉兴市/市本级';
const DEFAULT_DUPLICATE = '1';
const SUBJECT_DATA_PLACEHOLDER = '每行一个科目，数据用空格分隔。\n格式：科目名称 袋数总计 每箱袋数 尾箱超限袋数\n示例：\nCET4     191    60    10\n学考语文 122    30    3\n兼容Excel。可以在Excel中将科目的4列数据填写后，将数据复制到本输入框（不要复制标题）';
const SAMPLE_DATA = 'CET4\t191\t60\t10\nCET6\t122\t30\t5\n学考语文\t88\t40\t8\n日语\t156\t50\t10\n13000英语专升本\t95\t35\t5';

function createEmptySubject() {
    return {
        type: '',
        totalBags: '',
        bagsPerBox: '',
        maxExtraBags: '0'
    };
}

function createSubjectList() {
    return Array.from({ length: MAX_SUBJECTS }, createEmptySubject);
}

function toPositiveInt(value) {
    const parsed = Number.parseInt(value, 10);
    return Number.isFinite(parsed) && parsed > 0 ? parsed : 0;
}

function toNonNegativeInt(value) {
    const parsed = Number.parseInt(value, 10);
    return Number.isFinite(parsed) && parsed >= 0 ? parsed : 0;
}

function getFontSizeClass(type) {
    if (type.length > 12) {
        return 'font-shrink-3';
    }
    if (type.length > 8) {
        return 'font-shrink-2';
    }
    return '';
}

function calculateBoxCounts(totalBags, bagsPerBox, maxExtraBags) {
    if (totalBags <= 0 || bagsPerBox <= 0) {
        return [];
    }

    if (totalBags <= bagsPerBox) {
        return [totalBags];
    }

    const boxCounts = [];
    let remainingBags = totalBags;

    while (remainingBags > bagsPerBox) {
        boxCounts.push(bagsPerBox);
        remainingBags -= bagsPerBox;
    }

    if (remainingBags > 0) {
        const canMergeLastBox = boxCounts.length > 0 && remainingBags <= maxExtraBags;
        if (canMergeLastBox) {
            boxCounts[boxCounts.length - 1] += remainingBags;
        } else {
            boxCounts.push(remainingBags);
        }
    }

    return boxCounts;
}

function buildQrNumbers(enableQR, qrStart, qrEnd) {
    if (!enableQR || !qrStart || !qrEnd) {
        return [];
    }

    const startNum = Number.parseInt(qrStart, 10);
    const endNum = Number.parseInt(qrEnd, 10);
    if (!Number.isFinite(startNum) || !Number.isFinite(endNum) || endNum < startNum) {
        return [];
    }

    const length = Math.max(qrStart.length, qrEnd.length);
    return Array.from({ length: endNum - startNum + 1 }, (_, index) => String(startNum + index).padStart(length, '0'));
}

function normalizeSubject(subject) {
    return {
        type: subject.type.trim(),
        totalBags: toPositiveInt(subject.totalBags),
        bagsPerBox: toPositiveInt(subject.bagsPerBox),
        maxExtraBags: toNonNegativeInt(subject.maxExtraBags)
    };
}

function isSubjectValid(subject) {
    return Boolean(subject.type) && subject.totalBags > 0 && subject.bagsPerBox > 0;
}

function buildBoxes(snapshot) {
    const visibleSubjects = snapshot.subjects.slice(0, snapshot.subjectCount).map(normalizeSubject);
    const qrNumbers = buildQrNumbers(snapshot.enableQR, snapshot.qrStart, snapshot.qrEnd);
    const region = snapshot.region.trim() || '地区名称';

    let qrIndex = 0;
    const boxes = [];
    const showRightInfo = snapshot.showRightInfo;

    visibleSubjects.forEach((subject, subjectIndex) => {
        if (!isSubjectValid(subject)) {
            return;
        }

        const boxCounts = calculateBoxCounts(subject.totalBags, subject.bagsPerBox, subject.maxExtraBags);
        let cumulativeBags = 0;

        boxCounts.forEach((count, boxIndex) => {
            cumulativeBags += count;
            const qrData = snapshot.enableQR && qrIndex < qrNumbers.length
                ? `${snapshot.qrPrefix}${qrNumbers[qrIndex++]}`
                : '';

            boxes.push({
                key: `subject-${subjectIndex + 1}-box-${boxIndex + 1}`,
                region,
                boxNumber: `${boxIndex + 1}/${boxCounts.length}`,
                count,
                type: subject.type,
                totalBags: subject.totalBags,
                cumulativeBags,
                qrData,
                showRightInfo,
                fontSizeClass: getFontSizeClass(subject.type)
            });
        });
    });

    return {
        boxes,
        qrNumbers,
        insufficientQr: snapshot.enableQR && qrNumbers.length > 0 && boxes.length > qrNumbers.length
    };
}

function buildPages(boxes, duplicate) {
    if (duplicate) {
        return boxes.map(box => [box, box]);
    }

    const pages = [];
    for (let index = 0; index < boxes.length; index += 2) {
        pages.push(boxes.slice(index, index + 2));
    }
    return pages;
}

function parseManualLine(line) {
    const parts = line.trim().split(/\s+/).filter(Boolean);
    if (parts.length < 3) {
        return null;
    }

    const [type, totalBags, bagsPerBox, maxExtraBags = '0'] = parts;
    if (!type || Number.isNaN(Number.parseInt(totalBags, 10)) || Number.isNaN(Number.parseInt(bagsPerBox, 10))) {
        return null;
    }

    return {
        type,
        totalBags,
        bagsPerBox,
        maxExtraBags: Number.isNaN(Number.parseInt(maxExtraBags, 10)) ? '0' : maxExtraBags
    };
}

function cloneSnapshot(snapshot) {
    return {
        region: snapshot.region,
        duplicate: snapshot.duplicate,
        subjectCount: snapshot.subjectCount,
        enableQR: snapshot.enableQR,
        qrPrefix: snapshot.qrPrefix,
        qrStart: snapshot.qrStart,
        qrEnd: snapshot.qrEnd,
        showRightInfo: snapshot.showRightInfo,
        subjects: snapshot.subjects.map(subject => ({ ...subject }))
    };
}

function renderQrCode(element, text) {
    element.innerHTML = '';
    if (!text) {
        return;
    }

    new QRCode(element, {
        text,
        width: 80,
        height: 80,
        colorDark: '#000000',
        colorLight: '#ffffff',
        correctLevel: QRCode.CorrectLevel.H
    });
}

createApp({
    directives: {
        qrcode: {
            mounted(element, binding) {
                renderQrCode(element, binding.value);
            },
            updated(element, binding) {
                if (binding.value !== binding.oldValue) {
                    renderQrCode(element, binding.value);
                }
            }
        }
    },
    setup() {
        const region = ref(DEFAULT_REGION);
        const duplicate = ref(DEFAULT_DUPLICATE);
        const subjectCount = ref(1);
        const enableQR = ref(false);
        const showRightInfo = ref(true);
        const qrPrefix = ref('');
        const qrStart = ref('');
        const qrEnd = ref('');
        const manualInputVisible = ref(false);
        const manualText = ref('');
        const subjects = ref(createSubjectList());
        const lastQrWarningShown = ref(false);
        const subjectOptions = Array.from({ length: MAX_SUBJECTS }, (_, index) => index + 1);

        const buildFormSnapshot = () => ({
            region: region.value,
            duplicate: duplicate.value,
            subjectCount: Math.min(Math.max(subjectCount.value, 1), MAX_SUBJECTS),
            enableQR: enableQR.value,
            showRightInfo: showRightInfo.value,
            qrPrefix: qrPrefix.value,
            qrStart: qrStart.value,
            qrEnd: qrEnd.value,
            subjects: subjects.value.map(subject => ({ ...subject }))
        });

        const previewState = ref(cloneSnapshot(buildFormSnapshot()));

        const visibleSubjects = computed(() => subjects.value.slice(0, subjectCount.value).map(normalizeSubject));
        const hasValidSubjects = computed(() => visibleSubjects.value.some(isSubjectValid));
        const previewResult = computed(() => buildBoxes(previewState.value));
        const pages = computed(() => buildPages(previewResult.value.boxes, previewState.value.duplicate === '1'));

        let previewTimer = null;

        const flushPreview = () => {
            if (previewTimer) {
                clearTimeout(previewTimer);
                previewTimer = null;
            }

            previewState.value = cloneSnapshot(buildFormSnapshot());
        };

        const schedulePreview = () => {
            if (previewTimer) {
                clearTimeout(previewTimer);
            }

            previewTimer = setTimeout(() => {
                flushPreview();
            }, PREVIEW_DEBOUNCE_MS);
        };

        const subjectSignature = computed(() => subjects.value
            .map(subject => [subject.type, subject.totalBags, subject.bagsPerBox, subject.maxExtraBags].join('|'))
            .join('||'));

        watch(
            [region, duplicate, subjectCount, enableQR, showRightInfo, qrPrefix, qrStart, qrEnd, subjectSignature],
            schedulePreview,
            { immediate: true }
        );

        watch(
            () => previewResult.value.insufficientQr,
            insufficientQr => {
                if (insufficientQr && !lastQrWarningShown.value) {
                    console.warn('二维码数量不足以分配给所有箱子');
                    alert(`警告：生成的箱子总数(${previewResult.value.boxes.length})超过了二维码编号范围(${previewResult.value.qrNumbers.length})，部分箱子将没有二维码。`);
                }

                lastQrWarningShown.value = insufficientQr;
            },
            { immediate: true }
        );

        onBeforeUnmount(() => {
            if (previewTimer) {
                clearTimeout(previewTimer);
            }
        });

        const resetSubjects = () => {
            subjects.value = createSubjectList();
        };

        const toggleManualInput = () => {
            manualInputVisible.value = !manualInputVisible.value;
        };

        const generateLabels = () => {
            if (!hasValidSubjects.value) {
                alert('请至少填写一个有效的科目数据后再生成标签。');
                return;
            }

            flushPreview();
        };

        const fillSampleData = () => {
            if (manualText.value.trim() && !window.confirm('文本框中已有内容，是否清空并填充样例数据？')) {
                return;
            }

            manualText.value = SAMPLE_DATA;
        };

        const parseSubjectData = () => {
            const validData = manualText.value
                .split(/\r?\n/)
                .map(line => line.trim())
                .filter(Boolean)
                .map(parseManualLine)
                .filter(Boolean);

            if (validData.length === 0) {
                alert('没有找到有效的数据行！');
                return;
            }

            const limitedData = validData.slice(0, MAX_SUBJECTS);
            if (validData.length > MAX_SUBJECTS) {
                alert(`最多只支持导入 ${MAX_SUBJECTS} 个科目，已自动截取前 ${MAX_SUBJECTS} 行有效数据。`);
            }

            resetSubjects();
            subjectCount.value = limitedData.length;

            limitedData.forEach((subject, index) => {
                subjects.value[index] = {
                    type: subject.type,
                    totalBags: subject.totalBags,
                    bagsPerBox: subject.bagsPerBox,
                    maxExtraBags: subject.maxExtraBags
                };
            });

            flushPreview();
        };

        const printLabels = () => {
            window.print();
        };

        return {
            region,
            duplicate,
            subjectCount,
            subjectOptions,
            enableQR,
            showRightInfo,
            qrPrefix,
            qrStart,
            qrEnd,
            manualInputVisible,
            manualText,
            subjectDataPlaceholder: SUBJECT_DATA_PLACEHOLDER,
            subjects,
            pages,
            toggleManualInput,
            generateLabels,
            fillSampleData,
            parseSubjectData,
            printLabels
        };
    }
}).mount('#app');

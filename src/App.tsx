/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';

// =========================================================
// 1. 메뉴, 사이즈, 추가 옵션 데이터 및 타입 정의
// =========================================================

/** 음료 메뉴 항목 인터페이스 */
interface DrinkItem {
  id: string;
  name: string;
  price: number;
}

/** 사이즈 옵션 인터페이스 */
interface SizeItem {
  id: 'S' | 'M' | 'L';
  label: string;
  extraPrice: number;
}

/** 추가 옵션 인터페이스 */
interface ExtraOptionItem {
  id: string;
  name: string;
  extraPrice: number;
}

/** 접수된 주문 내역(게시판 아이템) 인터페이스 */
interface OrderRecord {
  orderNo: number;
  customerName: string;
  phone: string;
  drinkName: string;
  size: 'S' | 'M' | 'L';
  options: string[];
  quantity: number;
  requestNote: string;
  totalPrice: number;
  orderedAt: string;
}

/** 음료 선택 목록 (드롭다운용) */
const DRINK_MENU: DrinkItem[] = [
  { id: 'americano', name: '아메리카노', price: 3500 },
  { id: 'cafelatte', name: '카페라떼', price: 4000 },
  { id: 'cafemocha', name: '카페모카', price: 4500 },
  { id: 'vanillalatte', name: '바닐라라떼', price: 4500 },
  { id: 'greentealatte', name: '녹차라떼', price: 4500 },
];

/** 사이즈 선택 목록 (라디오 버튼용, 기본값: M) */
const SIZE_OPTIONS: SizeItem[] = [
  { id: 'S', label: 'S', extraPrice: 0 },
  { id: 'M', label: 'M', extraPrice: 500 },
  { id: 'L', label: 'L', extraPrice: 1000 },
];

/** 추가 옵션 목록 (체크박스용) */
const EXTRA_OPTIONS: ExtraOptionItem[] = [
  { id: 'shot', name: '샷 추가', extraPrice: 500 },
  { id: 'cream', name: '크림 추가', extraPrice: 500 },
  { id: 'syrup', name: '시럽 추가', extraPrice: 300 },
  { id: 'decaf', name: '디카페인', extraPrice: 0 },
];

export default function App() {
  // =========================================================
  // 2. 주문서 입력 상태(State) 관리
  // =========================================================
  const [customerName, setCustomerName] = useState<string>(''); // 1. 이름 (필수)
  const [phone, setPhone] = useState<string>(''); // 2. 전화번호
  const [selectedDrinkName, setSelectedDrinkName] = useState<string>(''); // 3. 음료 선택
  const [selectedSize, setSelectedSize] = useState<'S' | 'M' | 'L'>('M'); // 4. 사이즈 (기본 선택: M)
  const [selectedOptions, setSelectedOptions] = useState<string[]>([]); // 5. 추가 옵션 (복수 선택)
  const [quantity, setQuantity] = useState<number>(1); // 6. 수량 (1 ~ 10, 기본값: 1)
  const [requestNote, setRequestNote] = useState<string>(''); // 7. 요청사항

  // 알림 메시지(유효성 검사 경고) 및 주문 확인 메시지 상태
  const [validationAlert, setValidationAlert] = useState<string>('');
  const [confirmationMessage, setConfirmationMessage] = useState<string>('');
  const [excelSavedNotice, setExcelSavedNotice] = useState<string>('');

  // 카페 주문 게시판 목록 상태 (접수된 주문 저장)
  const [orderBoard, setOrderBoard] = useState<OrderRecord[]>([]);

  // 필수 입력칸 포커스 이동을 위한 Ref
  const nameInputRef = useRef<HTMLInputElement>(null);
  const drinkSelectRef = useRef<HTMLSelectElement>(null);

  // =========================================================
  // 3. 실시간 예상 금액 계산 로직
  // - 음료, 사이즈, 추가옵션, 수량이 바뀔 때마다 자동 재계산
  // =========================================================
  const estimatedPrice = useMemo(() => {
    // 음료가 아직 선택되지 않았으면 0원으로 표시
    const drinkObj = DRINK_MENU.find((d) => d.name === selectedDrinkName);
    if (!drinkObj) {
      return 0;
    }

    // 기본 음료 가격
    const baseDrinkPrice = drinkObj.price;

    // 선택한 사이즈 추가 금액
    const sizeObj = SIZE_OPTIONS.find((s) => s.id === selectedSize);
    const sizeExtraPrice = sizeObj ? sizeObj.extraPrice : 0;

    // 선택한 추가 옵션들의 금액 합산
    const optionsExtraPrice = selectedOptions.reduce((sum, optName) => {
      const foundOpt = EXTRA_OPTIONS.find((o) => o.name === optName);
      return sum + (foundOpt ? foundOpt.extraPrice : 0);
    }, 0);

    // 유효 수량 보정 (최소 1, 최대 10)
    const validQuantity = Math.min(10, Math.max(1, Number(quantity) || 1));

    // 최종 예상 금액 = (음료 가격 + 사이즈 추가금 + 옵션 추가금) * 수량
    return (baseDrinkPrice + sizeExtraPrice + optionsExtraPrice) * validQuantity;
  }, [selectedDrinkName, selectedSize, selectedOptions, quantity]);

  // =========================================================
  // 4. 이벤트 핸들러 (체크박스 토글, 수량 변경, 주문하기, 다시 작성, 엑셀 저장)
  // =========================================================

  /** 추가 옵션 체크박스 변경 처리 함수 */
  const handleOptionChange = (optionName: string) => {
    setSelectedOptions((prev) =>
      prev.includes(optionName)
        ? prev.filter((item) => item !== optionName)
        : [...prev, optionName]
    );
  };

  /** 수량 변경 처리 함수 (최소 1, 최대 10 제한) */
  const handleQuantityChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const rawValue = parseInt(e.target.value, 10);
    if (isNaN(rawValue)) {
      setQuantity(1);
      return;
    }
    const clampedValue = Math.min(10, Math.max(1, rawValue));
    setQuantity(clampedValue);
  };

  /**
   * 8. 주문하기 버튼 클릭 시 실행되는 함수
   * - 이름 미입력 시 "이름을 입력해주세요" 알림
   * - 음료 미선택 시 "음료를 선택해주세요" 알림
   * - 정상 입력 시 주문 확인 메시지 표시 및 주문 게시판에 추가
   */
  const handleOrderSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setExcelSavedNotice('');

    // 1) 이름 필수 입력 확인
    if (!customerName.trim()) {
      setValidationAlert('이름을 입력해주세요');
      setConfirmationMessage('');
      nameInputRef.current?.focus();
      return;
    }

    // 2) 음료 선택 여부 확인
    if (!selectedDrinkName) {
      setValidationAlert('음료를 선택해주세요');
      setConfirmationMessage('');
      drinkSelectRef.current?.focus();
      return;
    }

    // 유효성 통과 시 경고 메시지 초기화
    setValidationAlert('');

    // 추가 옵션 문자열 포맷팅 (예: " (샷 추가)" 또는 복수 선택 시 " (샷 추가, 시럽 추가)")
    const optionText =
      selectedOptions.length > 0 ? ` (${selectedOptions.join(', ')})` : '';

    // 천 단위 콤마가 적용된 총 금액 문자열
    const formattedTotal = estimatedPrice.toLocaleString('ko-KR');

    // 주문 확인 메시지 생성
    // 예: "홍길동님, 카페라떼 M사이즈 (샷 추가) 1잔, 총 5,000원 주문이 접수되었습니다!"
    const message = `${customerName.trim()}님, ${selectedDrinkName} ${selectedSize}사이즈${optionText} ${quantity}잔, 총 ${formattedTotal}원 주문이 접수되었습니다!`;
    setConfirmationMessage(message);

    // 카페 주문 게시판 목록에 새 주문 데이터 추가
    const now = new Date();
    const formattedTime = now.toLocaleString('ko-KR', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    });

    const newRecord: OrderRecord = {
      orderNo: orderBoard.length + 1,
      customerName: customerName.trim(),
      phone: phone.trim() || '-',
      drinkName: selectedDrinkName,
      size: selectedSize,
      options: [...selectedOptions],
      quantity,
      requestNote: requestNote.trim() || '-',
      totalPrice: estimatedPrice,
      orderedAt: formattedTime,
    };

    setOrderBoard((prev) => [newRecord, ...prev]);
  };

  /**
   * 9. 다시 작성 버튼 클릭 시 실행되는 함수
   * - 모든 입력값, 선택 상태, 예상 금액 및 메시지를 초기 상태로 리셋
   */
  const handleReset = () => {
    setCustomerName('');
    setPhone('');
    setSelectedDrinkName('');
    setSelectedSize('M'); // 기본 선택 M으로 복원
    setSelectedOptions([]);
    setQuantity(1); // 기본값 1로 복원
    setRequestNote('');
    setValidationAlert('');
    setConfirmationMessage('');
    setExcelSavedNotice('');
  };

  /**
   * 10. 주문서의 양식을 엑셀 파일(.xlsx)로서 데이터를 저장하는 함수
   * - 시트 1: 현재 작성 중인 '숭실 카페 주문서 양식' 및 입력 데이터
   * - 시트 2: 누적된 '주문 접수 게시판 내역' 데이터
   */
  const handleDownloadExcel = () => {
    setValidationAlert('');

    const now = new Date();
    const dateStr = now.toLocaleString('ko-KR');

    // [시트 1] 주문서 양식 및 현재 입력 데이터 구성 (2차원 배열 형태)
    const formSheetData = [
      ['☕ 숭실 카페 (Soongsil Cafe) - 주문서 양식'],
      ['부제', '당신의 하루에 바이브를 더하다'],
      ['파일 생성 일시', dateStr],
      [],
      ['[1. 주문서 작성 데이터]'],
      ['항목', '입력 데이터', '비고 / 선택 가능 옵션 안내'],
      ['1. 이름 (필수)', customerName.trim() || '(미입력)', '주문자 성함 (필수 입력)'],
      ['2. 전화번호', phone.trim() || '(미입력)', '연락처 (예: 010-1234-5678)'],
      [
        '3. 음료 선택',
        selectedDrinkName || '(미선택)',
        '아메리카노(3,500원) / 카페라떼(4,000원) / 카페모카(4,500원) / 바닐라라떼(4,500원) / 녹차라떼(4,500원)',
      ],
      [
        '4. 사이즈',
        `${selectedSize} 사이즈`,
        'S(+0원) / M(+500원, 기본) / L(+1,000원)',
      ],
      [
        '5. 추가 옵션',
        selectedOptions.length > 0 ? selectedOptions.join(', ') : '선택 없음',
        '샷 추가(+500원) / 크림 추가(+500원) / 시럽 추가(+300원) / 디카페인(+0원)',
      ],
      ['6. 수량', `${quantity}잔`, '최소 1잔 ~ 최대 10잔 (기본값 1)'],
      ['7. 요청사항', requestNote.trim() || '(없음)', '얼음 양, 농도, 포장 여부 등 자유 기재'],
      ['8. 예상 금액 (원)', `${estimatedPrice.toLocaleString('ko-KR')}원`, '실시간 자동 계산 금액'],
      [],
      ['[2. 숭실 카페 메뉴 및 옵션 단가표]'],
      ['구분', '메뉴 / 옵션명', '단가 (원)'],
      ...DRINK_MENU.map((d) => ['음료', d.name, `${d.price.toLocaleString('ko-KR')}원`]),
      ...SIZE_OPTIONS.map((s) => [
        '사이즈',
        `${s.label} 사이즈`,
        `+${s.extraPrice.toLocaleString('ko-KR')}원`,
      ]),
      ...EXTRA_OPTIONS.map((o) => [
        '추가 옵션',
        o.name,
        `+${o.extraPrice.toLocaleString('ko-KR')}원`,
      ]),
    ];

    // [시트 2] 주문 게시판 누적 데이터 구성
    // 접수된 주문이 있으면 해당 목록을 저장하고, 아직 접수 전이면 현재 작성 중인 데이터를 1행 샘플로 함께 구성
    const boardRows =
      orderBoard.length > 0
        ? orderBoard.map((order) => ({
            주문번호: order.orderNo,
            주문일시: order.orderedAt,
            이름: order.customerName,
            전화번호: order.phone,
            선택음료: order.drinkName,
            사이즈: order.size,
            추가옵션: order.options.length > 0 ? order.options.join(', ') : '없음',
            수량: `${order.quantity}잔`,
            요청사항: order.requestNote,
            '결제금액(원)': order.totalPrice,
            '결제금액(표시)': `${order.totalPrice.toLocaleString('ko-KR')}원`,
          }))
        : [
            {
              주문번호: 1,
              주문일시: dateStr,
              이름: customerName.trim() || '(작성중)',
              전화번호: phone.trim() || '-',
              선택음료: selectedDrinkName || '(미선택)',
              사이즈: selectedSize,
              추가옵션:
                selectedOptions.length > 0 ? selectedOptions.join(', ') : '없음',
              수량: `${quantity}잔`,
              요청사항: requestNote.trim() || '-',
              '결제금액(원)': estimatedPrice,
              '결제금액(표시)': `${estimatedPrice.toLocaleString('ko-KR')}원`,
            },
          ];

    // 엑셀 워크북(Workbook) 생성
    const workbook = XLSX.utils.book_new();

    // 시트 1 생성 및 열 너비 지정
    const worksheetForm = XLSX.utils.aoa_to_sheet(formSheetData);
    worksheetForm['!cols'] = [
      { wch: 20 }, // A열: 항목명
      { wch: 28 }, // B열: 입력 데이터
      { wch: 62 }, // C열: 비고 / 옵션 안내
    ];
    XLSX.utils.book_append_sheet(workbook, worksheetForm, '주문서_양식');

    // 시트 2 생성 및 열 너비 지정
    const worksheetBoard = XLSX.utils.json_to_sheet(boardRows);
    worksheetBoard['!cols'] = [
      { wch: 10 }, // 주문번호
      { wch: 22 }, // 주문일시
      { wch: 14 }, // 이름
      { wch: 16 }, // 전화번호
      { wch: 16 }, // 선택음료
      { wch: 10 }, // 사이즈
      { wch: 26 }, // 추가옵션
      { wch: 10 }, // 수량
      { wch: 28 }, // 요청사항
      { wch: 14 }, // 결제금액(원)
      { wch: 16 }, // 결제금액(표시)
    ];
    XLSX.utils.book_append_sheet(workbook, worksheetBoard, '주문_게시판_데이터');

    // 엑셀 파일(.xlsx) 다운로드 실행
    const fileName = `숭실카페_주문서_데이터.xlsx`;
    XLSX.writeFile(workbook, fileName);

    setExcelSavedNotice(`'${fileName}' 엑셀 파일이 성공적으로 저장되었습니다.`);
  };

  return (
    <div className="min-h-screen bg-[#f0f7fc] py-8 px-4">
      {/* =========================================================
       * 전체 컨테이너
       * - 최대 너비 520px, 가운데 정렬
       * - 둥근 모서리, 부드러운 그림자, 연푸른색조 카페 디자인
       * ========================================================= */}
      <main className="max-w-[520px] mx-auto bg-[#f9fcff] rounded-2xl shadow-[0_10px_30px_rgba(43,108,176,0.12)] border border-[#cfe4f5] p-6 sm:p-8">
        {/* =========================================================
         * [페이지 상단]
         * - 카페 로고: ☕ 이모지 크게
         * - 카페 이름: "숭실 카페"
         * - 부제: "당신의 하루에 바이브를 더하다"
         * ========================================================= */}
        <header className="text-center pb-6 mb-6 border-b border-[#d9eaf8]">
          <div
            className="text-6xl mb-3 select-none leading-none"
            role="img"
            aria-label="카페 로고"
          >
            ☕
          </div>
          <h1 className="text-2xl sm:text-3xl font-bold text-[#2b6cb0] tracking-tight mb-1.5">
            숭실 카페
          </h1>
          <p className="text-sm sm:text-base text-[#4a76a8]">
            당신의 하루에 바이브를 더하다
          </p>
        </header>

        {/* =========================================================
         * [주문서 폼 영역]
         * - 모든 input/select/textarea에 label(htmlFor) 1:1 연결
         * ========================================================= */}
        <form onSubmit={handleOrderSubmit} noValidate className="space-y-5">
          {/* 1. 이름 (필수, text) */}
          <div>
            <label
              htmlFor="customer-name"
              className="block text-sm font-bold text-[#2b6cb0] mb-1.5"
            >
              1. 이름 <span className="text-[#e53e3e]">*</span>
            </label>
            <input
              ref={nameInputRef}
              id="customer-name"
              type="text"
              required
              value={customerName}
              onChange={(e) => {
                setCustomerName(e.target.value);
                if (validationAlert) setValidationAlert('');
              }}
              placeholder="주문자 이름을 입력해주세요 (예: 홍길동)"
              className="cafe-input"
            />
          </div>

          {/* 2. 전화번호 (tel) */}
          <div>
            <label
              htmlFor="customer-phone"
              className="block text-sm font-bold text-[#2b6cb0] mb-1.5"
            >
              2. 전화번호
            </label>
            <input
              id="customer-phone"
              type="tel"
              value={phone}
              onChange={(e) => setPhone(e.target.value)}
              placeholder="010-0000-0000"
              className="cafe-input"
            />
          </div>

          {/* 3. 음료 선택 (드롭다운) */}
          <div>
            <label
              htmlFor="drink-select"
              className="block text-sm font-bold text-[#2b6cb0] mb-1.5"
            >
              3. 음료 선택 <span className="text-[#e53e3e]">*</span>
            </label>
            <select
              ref={drinkSelectRef}
              id="drink-select"
              value={selectedDrinkName}
              onChange={(e) => {
                setSelectedDrinkName(e.target.value);
                if (validationAlert) setValidationAlert('');
              }}
              className="cafe-input cursor-pointer"
            >
              <option value="">-- 음료를 선택해주세요 --</option>
              {DRINK_MENU.map((drink) => (
                <option key={drink.id} value={drink.name}>
                  {drink.name} {drink.price.toLocaleString('ko-KR')}원
                </option>
              ))}
            </select>
          </div>

          {/* 4. 사이즈 (라디오 버튼, 가로 배치, M 기본 선택) */}
          <fieldset>
            <legend className="block text-sm font-bold text-[#2b6cb0] mb-2">
              4. 사이즈
            </legend>
            <div className="flex flex-wrap items-center gap-4 bg-[#eef6fc] p-3 rounded-[8px] border border-[#cfe4f5]">
              {SIZE_OPTIONS.map((sizeItem) => {
                const radioId = `size-radio-${sizeItem.id}`;
                return (
                  <div key={sizeItem.id} className="flex items-center gap-1.5">
                    <input
                      id={radioId}
                      type="radio"
                      name="drink-size"
                      value={sizeItem.id}
                      checked={selectedSize === sizeItem.id}
                      onChange={() => setSelectedSize(sizeItem.id)}
                      className="cafe-choice-input"
                    />
                    <label
                      htmlFor={radioId}
                      className="text-sm text-[#1a365d] cursor-pointer select-none font-medium"
                    >
                      {sizeItem.label} +{sizeItem.extraPrice.toLocaleString('ko-KR')}원
                    </label>
                  </div>
                );
              })}
            </div>
          </fieldset>

          {/* 5. 추가 옵션 (체크박스, 가로 배치) */}
          <fieldset>
            <legend className="block text-sm font-bold text-[#2b6cb0] mb-2">
              5. 추가 옵션
            </legend>
            <div className="flex flex-wrap items-center gap-4 bg-[#eef6fc] p-3 rounded-[8px] border border-[#cfe4f5]">
              {EXTRA_OPTIONS.map((option) => {
                const checkboxId = `extra-option-${option.id}`;
                const isChecked = selectedOptions.includes(option.name);
                return (
                  <div key={option.id} className="flex items-center gap-1.5">
                    <input
                      id={checkboxId}
                      type="checkbox"
                      checked={isChecked}
                      onChange={() => handleOptionChange(option.name)}
                      className="cafe-choice-input"
                    />
                    <label
                      htmlFor={checkboxId}
                      className="text-sm text-[#1a365d] cursor-pointer select-none font-medium"
                    >
                      {option.name} +{option.extraPrice.toLocaleString('ko-KR')}원
                    </label>
                  </div>
                );
              })}
            </div>
          </fieldset>

          {/* 6. 수량 (number 타입, 최소 1, 최대 10, 기본값 1) */}
          <div>
            <label
              htmlFor="order-quantity"
              className="block text-sm font-bold text-[#2b6cb0] mb-1.5"
            >
              6. 수량 (1 ~ 10잔)
            </label>
            <input
              id="order-quantity"
              type="number"
              min={1}
              max={10}
              value={quantity}
              onChange={handleQuantityChange}
              className="cafe-input tabular-nums"
            />
          </div>

          {/* 7. 요청사항 (textarea) */}
          <div>
            <label
              htmlFor="order-request"
              className="block text-sm font-bold text-[#2b6cb0] mb-1.5"
            >
              7. 요청사항
            </label>
            <textarea
              id="order-request"
              rows={3}
              value={requestNote}
              onChange={(e) => setRequestNote(e.target.value)}
              placeholder="얼음 적게, 시럽 별도 등 요청사항을 입력해주세요"
              className="cafe-input resize-y"
            />
          </div>

          {/* =========================================================
           * [실시간 예상 금액 표시 영역]
           * - 주문하기 버튼 바로 위에 큰 글씨(24px), 푸른색(#2b6cb0), 굵게, 가운데 정렬
           * - 천 단위 콤마(toLocaleString) 적용
           * ========================================================= */}
          <div className="pt-2 pb-1 text-center">
            <p
              className="text-[24px] font-bold text-[#2b6cb0] tabular-nums"
              aria-live="polite"
            >
              예상 금액: {estimatedPrice.toLocaleString('ko-KR')}원
            </p>
          </div>

          {/* 유효성 검사 알림 메시지 (이름 미입력 / 음료 미선택 시 표시) */}
          {validationAlert && (
            <div
              role="alert"
              className="p-3.5 rounded-[8px] bg-[#fdecea] border border-[#f5c2c0] text-[#b71c1c] text-sm font-bold text-center"
            >
              ⚠️ {validationAlert}
            </div>
          )}

          {/* =========================================================
           * [버튼 영역]
           * 8. 주문하기 버튼: 푸른색 배경(#2b6cb0), 흰색 글씨, hover시 약간 밝게
           * 9. 다시 작성 버튼: 모든 입력과 금액 초기화
           * 10. 주문서의 양식을 엑셀파일로서 데이터를 저장할 수 있게하는 버튼
           * ========================================================= */}
          <div className="space-y-2.5 pt-1">
            <div className="flex gap-2.5">
              {/* 8. 주문하기 버튼 */}
              <button
                type="submit"
                className="flex-1 py-3 px-4 rounded-[8px] bg-[#2b6cb0] hover:bg-[#3b82ce] active:bg-[#1e4e8c] text-white font-bold text-base transition-colors cursor-pointer whitespace-nowrap shadow-sm"
              >
                주문하기
              </button>

              {/* 9. 다시 작성 버튼 */}
              <button
                type="button"
                onClick={handleReset}
                className="py-3 px-5 rounded-[8px] bg-[#e1effa] hover:bg-[#cfe4f5] text-[#2b6cb0] font-bold text-base transition-colors cursor-pointer whitespace-nowrap"
              >
                다시 작성
              </button>
            </div>

            {/* 10. 주문서 양식 및 데이터 엑셀(.xlsx) 저장 버튼 */}
            <button
              type="button"
              onClick={handleDownloadExcel}
              className="w-full py-2.5 px-4 rounded-[8px] bg-[#eaf4fc] hover:bg-[#d8eaf9] border border-[#b8d8f3] text-[#2b6cb0] font-bold text-sm transition-colors cursor-pointer flex items-center justify-center gap-2"
            >
              <span>📊</span>
              <span>주문서 양식 및 데이터 엑셀파일(.xlsx) 저장</span>
            </button>
          </div>
        </form>

        {/* 엑셀 파일 저장 완료 안내 메시지 */}
        {excelSavedNotice && (
          <div className="mt-4 p-3 rounded-[8px] bg-[#e6f2fc] text-[#2b6cb0] text-xs font-medium text-center">
            {excelSavedNotice}
          </div>
        )}

        {/* =========================================================
         * [주문 확인 메시지]
         * - 연두색 배경, 초록 글씨, 둥근 모서리
         * - 예: "홍길동님, 카페라떼 M사이즈 (샷 추가) 1잔, 총 5,000원 주문이 접수되었습니다!"
         * ========================================================= */}
        {confirmationMessage && (
          <div
            role="status"
            className="mt-5 p-4 rounded-[8px] bg-[#e6f4ea] border border-[#b7e1cd] text-[#1e6f3d] font-bold text-sm sm:text-base text-center leading-relaxed"
          >
            {confirmationMessage}
          </div>
        )}

        {/* =========================================================
         * [카페 주문 게시판 (접수된 주문 내역)]
         * - 주문이 접수되면 하단 게시판에 실시간으로 추가되어 내역 확인 가능
         * ========================================================= */}
        {orderBoard.length > 0 && (
          <section className="mt-8 pt-6 border-t border-[#d9eaf8]">
            <div className="flex items-center justify-between mb-3">
              <h2 className="text-base font-bold text-[#2b6cb0]">
                📋 접수된 주문 게시판 ({orderBoard.length}건)
              </h2>
              <span className="text-xs text-[#4a76a8] tabular-nums">
                총 누적 금액:{' '}
                {orderBoard
                  .reduce((acc, cur) => acc + cur.totalPrice, 0)
                  .toLocaleString('ko-KR')}
                원
              </span>
            </div>

            <div className="space-y-2.5 max-h-64 overflow-y-auto pr-1">
              {orderBoard.map((item) => (
                <article
                  key={item.orderNo}
                  className="p-3.5 rounded-[8px] bg-[#eef6fc] border border-[#cfe4f5] text-sm"
                >
                  <div className="flex items-center justify-between mb-1">
                    <span className="font-bold text-[#2b6cb0]">
                      #{item.orderNo} {item.customerName}님
                    </span>
                    <span className="font-bold text-[#2b6cb0] tabular-nums">
                      {item.totalPrice.toLocaleString('ko-KR')}원
                    </span>
                  </div>
                  <p className="text-[#1a365d]">
                    {item.drinkName} · {item.size}사이즈
                    {item.options.length > 0
                      ? ` · ${item.options.join(', ')}`
                      : ''}{' '}
                    · {item.quantity}잔
                  </p>
                  {(item.phone !== '-' || item.requestNote !== '-') && (
                    <p className="text-xs text-[#4a76a8] mt-1">
                      {item.phone !== '-' && `연락처: ${item.phone}`}
                      {item.phone !== '-' && item.requestNote !== '-' && ' · '}
                      {item.requestNote !== '-' && `요청: ${item.requestNote}`}
                    </p>
                  )}
                </article>
              ))}
            </div>
          </section>
        )}
      </main>
    </div>
  );
}

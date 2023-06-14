# LyricsPPTMaker
* [SundayWorshipPPTMaker](https://github.com/farka-k/SundayWorshipPPTMaker)에서 가사검색/입력 기능만 분리하여 범용적으로 사용
* MVVM패턴 실습
-----------------------------------------------------
## Layout & Usage
![image](https://github.com/farka-k/LyricsPPTMaker/assets/32349691/9200be81-c077-4c08-9155-9c690179557e)
### 1. 곡 제목을 입력 후 버튼을 누르면 곡이 검색된다.
### 2. 검색된 곡의 제목, 아티스트, 앨범, 가사가 표시된다.
   * 화살표 버튼으로 표시 곡을 변경
   * 구글 검색 API로 부터 상위 5개만을 가져오므로 원하는 곡과 다른 곡이 결과에 포함될 수 있음
   * 가사는 수정이 가능하며, 한 문단이 하나에 슬라이드에 삽입된다.
   
     ![image](https://github.com/farka-k/LyricsPPTMaker/assets/32349691/690ecc29-98fb-4aeb-876e-49a954bbf140)
   * 수정 가사는 저장되지 않으므로 표시 곡 변경시 주의

### 3. 슬라이드 스타일에 대한 프리셋 설정(관련 기능 미구현)
### 4. 화면크기, 배경색, 글꼴색, 글꼴, 글꼴크기, 텍스트효과, 텍스트위치 설정
   *  Offset값이 양수면 기준 위치 아래로, 음수면 위로 이동
   *  VerticalAlignment 설정에 따라 Offset 범위 변동
   
### 5. 가사 슬라이드에 곡 제목 표시 여부(관련 기능 미구현)
### 6. 모든 입력 요소를 초기값으로 설정
### 7. 현재 설정으로 가사 슬라이드를 클립보드에 복사.    
   * 프레젠테이션에서 '원본 서식 유지'로 붙여넣기 할 것

### 8. 현재 스타일에 대한 미리보기 팝업을 표시한다.
   * 가사 검색시 On
   * On상태에서 버튼 클릭시, 혹은 팝업 클릭시 Off

## Next Update
> 1. Key Binding
> 2. Preset
> 3. 제목 슬라이드 추가, 가사 슬라이드에 제목 표시 추가
> 4. 가사 문단 자동 분리

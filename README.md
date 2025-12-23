# mermaid to pptx

A project that converts Mermaid diagrams into PowerPoint (.pptx) shapes by rendering Mermaid to SVG and then analyzing
and reconstructing the diagram using native PowerPoint objects.

This project focuses on vector-based conversion, not image embedding, so the resulting slides remain fully editable and
scalable.

---

## Features

- Accepts Mermaid diagram syntax as input
- Renders Mermaid diagrams to SVG
- Parses SVG structure (shapes, lines, text)
- Converts elements into PowerPoint (.pptx) shapes
- Produces fully editable, vector-based slides

---

## Sample Output

```mermaid
sequenceDiagram
%% title Sample Sequence Diagram (Abstract)
    autonumber
    participant User as User
    participant App as Client
    participant API as API Server

    box rgb(235,245,235) "Internal Area"
        participant SVC as Service
        participant DB@{ "type": "database", "as": "DB" }
        participant CACHE@{ "type": "database", "as": "Cache" }
    end

%% -------------------------
%% 데이터 조회
%% -------------------------
    note over User, API: Data 조회
    User ->> App: 조회 요청
    App ->> API: 조회 요청 전달
    API ->> SVC: 조회 처리 요청
    SVC ->> DB: 데이터 조회
    DB -->> SVC: 조회 결과
    SVC -->> API: 결과 반환
    API -->> App: 응답 전달
    App -->> User: 화면 표시
%% -------------------------
%% 등록 처리
%% -------------------------
    note over User, API: 등록 처리
    User ->> App: 등록 요청
    App ->> API: 등록 요청 전달
    API ->> SVC: 등록 처리 요청
    SVC ->> DB: 유효성 확인

    alt 조건 만족
        SVC ->> DB: 데이터 등록
        SVC -->> API: 등록 성공
    else 조건 불만족
        SVC -->> API: 등록 실패
    end

    API -->> App: 등록 결과 응답
    App -->> User: 결과 안내
%% -------------------------
%% 세션 생성
%% -------------------------
    note over User, API: 세션 생성
    User ->> App: 로그인 요청
    App ->> API: 인증 요청
    API ->> SVC: 인증 처리
    SVC ->> CACHE: 세션 생성
    CACHE -->> SVC: 세션 Token
    SVC -->> API: Token 반환
    API -->> App: 로그인 성공
    App -->> User: 로그인 완료
%% -------------------------
%% 세션 기반 요청
%% -------------------------
    note over User, API: 세션 기반 요청
    App ->> API: 요청 (Token)
    API ->> CACHE: Token 검증
    CACHE -->> API: 사용자 식별자
    API ->> DB: 데이터 조회
    DB -->> API: 결과
    API -->> App: 응답 반환
```

![SVG Result](sample/sample_pptx.png)

---

## Overall Flow

```text
Mermaid Code
     ↓
Mermaid Renderer (SVG)
     ↓
SVG Parsing
     ↓
Shape & Text Extraction
     ↓
PowerPoint (.pptx)

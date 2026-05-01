import pandas as pd
import numpy as np
import os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

def run_analysis_to_excel():
    # 경로 설정
    file_path = 'data/products_dataset.csv'
    img_path = 'images/image_668a63.png'  # 업로드하신 통계 표 이미지
    output_dir = 'output'
    excel_path = f'{output_dir}/product_analysis_report.xlsx'

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        # 1. 데이터 로드
        df = pd.read_csv(file_path)
        print("✅ 데이터를 성공적으로 불러왔습니다.")

        # 2. 수치형 데이터 통계 (이미지의 모든 항목: count, mean, std, min, 25%, 50%, 75%, max)
        # .describe().T를 사용하면 이미지와 동일한 구성의 통계표가 생성됩니다.
        stats = df.describe().T
        stats = stats[['count', 'mean', 'std', 'min', '25%', '50%', '75%', 'max']]

        # 3. 카테고리 별 수치형 데이터 평균 추출
        category_avg = df.groupby('product_category_name').mean(numeric_only=True)

        # 4. 데이터 정보 (컬럼명, 타입, Null 여부)
        info_df = pd.DataFrame({
            '컬럼명': df.columns,
            '데이터 타입': df.dtypes.values,
            '결측치 아닌 개수': df.count().values,
            'Null 존재 여부': df.isnull().any().values
        })

        # 5. 엑셀 파일 생성 및 시트별 저장
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            info_df.to_excel(writer, sheet_name='1.데이터_정보', index=False)
            stats.to_excel(writer, sheet_name='2.수치형_통계_상세')
            category_avg.to_excel(writer, sheet_name='3.카테고리별_평균')
            df.head(100).to_excel(writer, sheet_name='4.데이터_샘플(100개)', index=False)

        # 6. 엑셀에 이미지 삽입 (요구사항)
        if os.path.exists(img_path):
            wb = load_workbook(excel_path)
            ws = wb['2.수치형_통계_상세']

            # 이미지 객체 생성 및 크기 조정
            img = Image(img_path)
            img.width = 450  # 엑셀 내 이미지 너비 조절
            img.height = 300 # 엑셀 내 이미지 높이 조절

            # 통계표 옆(J2 셀)에 이미지 붙이기
            ws.add_image(img, 'J2')
            wb.save(excel_path)
            print("📸 엑셀 시트에 통계 이미지가 삽입되었습니다.")

        print(f"🎉 모든 요구사항이 반영된 엑셀 보고서가 생성되었습니다: {excel_path}")

    except FileNotFoundError:
        print(f"❌ 에러: 파일을 찾을 수 없습니다. 경로를 확인하세요.")
    except Exception as e:
        print(f"❌ 에러 발생: {e}")

if __name__ == "__main__":
    run_analysis_to_excel()
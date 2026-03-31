/home/tomo/work_office_edinet_api/富士電機/有報_富士電機_202503.xlsx
で検証しました。

報告年度の前年度のデータを使うという、古いロジックが残ってしまっています。
具体的には、連結_セグメント情報等(日本基準)_分析_PPM分析用　シートの　報告年度：2023-03-31、会計年度：2023-03-31　に、エネルギーセグメントの売上が記載されています。
　売上	2023-03-31	当期	2023-03-31	333,272
これは、報告年度、2024-03-31 の前期のデータから取得した値です。

正しくは、報告年度、2023-03-31には、エネルギーセグメントは存在しません。

報告年度の前期、当期のセグメントは、必ず、ペアになっている必要があります。
富士電機の例だと、報告年度の、2023-03-31　、2021-03-31、2019-03-31,2017-03-31がペアになっていません。

問題を理解しました。analysis_wsはマージされたデータを持っており、2024年の前期データが2023-03-31の期間に「エネルギー」セグメントとして混入しています。各報告年度のZIPに実際に存在したセグメントのみを使うようにフィルタリングが必要です。

まず convert_xbrl_to_excel.py の filing_pairs に current_dims/prior_dims を追加します：

Edit convert_xbrl_to_excel.py
Added 19 lines

        # --- 各ZIPの当期/前期ペアを構築（古→新の昇順）---
        # results は新しい年度順ソート済みなので reversed で古→新の順に走査する
        # current_dims / prior_dims: そのZIPの当期・前期に実在したセグメントdim名の集合
        # これをPPMシートでのフィルタリングに使い、他のZIPからのデータ混入を防ぐ
        filing_pairs = []
        for res in reversed(results):
            seg_periods = set()
            seg_dim_by_period = {}   # period_str -> set of dim_labels
            for el, el_vals in res.get('values', {}).items():
                if el == '_metadata':
                    continue
                for (std, dim, period) in el_vals.keys():
                    if _is_segment_dim(dim):
                        seg_periods.add(period)
                        if period not in seg_dim_by_period:
                            seg_dim_by_period[period] = set()
                        seg_dim_by_period[period].add(dim)
            sorted_p = sorted(seg_periods)
            if len(sorted_p) >= 2:
                cur_p = sorted_p[-1]
                pri_p = sorted_p[-2]
                filing_pairs.append({
                    'current': cur_p,
                    'prior': pri_p,
                    'current_dims': seg_dim_by_period.get(cur_p, set()),
                    'prior_dims':   seg_dim_by_period.get(pri_p, set()),
                })
            elif len(sorted_p) == 1:
                cur_p = sorted_p[0]
                filing_pairs.append({
                    'current': cur_p,
                    'prior': None,
                    'current_dims': seg_dim_by_period.get(cur_p, set()),
                    'prior_dims':   set(),
                })


PPM作成用の集約表は、C列から開始する。
PPM作成用の集約表の、書式設定は、変更前と同様
軸の最大、最小のロジックも　変更前と同様

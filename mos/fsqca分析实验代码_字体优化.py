"""
fsQCA分析实验代码
功能：实现公共图书馆AI虚拟馆员用户间歇性中辍行为的完整分析流程
包括数据处理、变量定义、校准、必要性分析、真值表生成与组态分析
基于模拟问卷数据，结合扎根理论与fsQCA方法
"""

import pandas as pd
import numpy as np
from scipy import stats
import warnings

# 忽略警告信息
warnings.filterwarnings('ignore')

import matplotlib
import matplotlib.pyplot as plt
import seaborn as sns

# Mac系统中文字体配置
matplotlib.rcParams['font.sans-serif'] = ['PingFang SC', 'Hiragino Sans GB', 'Arial Unicode MS', 'DejaVu Sans']
matplotlib.rcParams['axes.unicode_minus'] = False
try:
    import seaborn as sns
    sns.set(font='PingFang SC')
except:
    pass
print("已配置Mac系统中文字体支持")

class FsQCAAnalyzer:
    """
    fsQCA（模糊集定性比较分析）分析器
    实现从原始数据到组态分析的完整流程
    """

    def __init__(self, data_path):
        """
        初始化分析器
        
        Args:
            data_path (str): 数据文件路径
        """
        self.data_path = data_path
        self.raw_data = None
        self.processed_data = None
        self.calibrated_data = {}
        self.results = {
            'variable_definition': {},
            'calibration_anchors': {},
            'necessity_analysis': {},
            'truth_table': pd.DataFrame(),
            'configurations': []
        }

    def load_data(self):
        """
        加载原始数据
        
        Raises:
            FileNotFoundError: 当数据文件不存在时
            Exception: 当数据读取失败时
        """
        try:
            # 读取Excel文件第一个工作表
            self.raw_data = pd.read_excel(self.data_path, sheet_name='Sheet1')
            print(f"成功加载数据：{self.raw_data.shape[0]} 行，{self.raw_data.shape[1]} 列")
            
            # 验证关键字段是否存在
            required_columns = ['是否再次使用', 'A1', 'B1', 'C1', 'D1']
            missing_cols = [col for col in required_columns if col not in self.raw_data.columns]
            if missing_cols:
                raise ValueError(f"缺失必要列: {missing_cols}")
                
        except FileNotFoundError:
            raise FileNotFoundError(f"未找到数据文件: {self.data_path}")
        except Exception as e:
            raise Exception(f"读取数据文件时出错: {str(e)}")

    def define_variables(self):
        """
        基于扎根理论结果定义分析变量
        将原始测量项聚合为理论构念
        """
        df = self.raw_data.copy()
        
        # 定义变量映射关系（基于扎根理论三级编码）
        variable_mapping = {
            '认知负荷过高': ['A1', 'A2', 'A3'],
            '情感疏离感': ['A4', 'A5'],
            '隐私侵犯担忧': ['A6', 'A7'],
            '成本沉没感知': ['A8'],
            '技术响应失效': ['B1', 'B2', 'B3'],
            '功能局限性': ['B4', 'B5'],
            '信息可信度质疑': ['B6', 'B7'],
            '交互界面障碍': ['B8', 'B9'],
            '替代渠道便利': ['C1', 'C2', 'C3'],
            '使用动机减弱': ['C4', 'C5'],
            '外部干扰因素': ['C6', 'C7'],
            '社会规范影响': ['D1', 'D2', 'D3'],
            '再启用触发事件': ['D4', 'D5', 'D6'],
            '功能升级感知': ['D7'],
            '情感依恋重建': ['D8', 'D9']
        }
        
        # 计算复合变量均值
        processed_df = df[['年龄段', '性别', '教育程度', '职业类型', '是否再次使用']].copy()
        
        for construct, items in variable_mapping.items():
            processed_df[construct] = df[items].mean(axis=1)
            self.results['variable_definition'][construct] = {
                'type': 'condition',
                'source_items': items,
                'measurement': 'mean'
            }
        
        # 添加结果变量
        processed_df['用户间歇性中辍后再次使用'] = df['是否再次使用'].map({'是': 1, '否': 0})
        self.results['variable_definition']['用户间歇性中辍后再次使用'] = {
            'type': 'outcome',
            'source_column': '是否再次使用'
        }
        
        self.processed_data = processed_df
        print("变量定义完成，共创建15个条件变量和1个结果变量")

    def calibrate_sets(self):
        """
        模糊集校准
        使用直接校准法，基于四分位数确定锚点
        """
        if self.processed_data is None:
            raise ValueError("请先执行define_variables()")

        df = self.processed_data.drop(columns=['年龄段', '性别', '教育程度', '职业类型'])
        calibrated_df = df[['用户间歇性中辍后再次使用']].copy()

        for col in df.columns:
            if col == '用户间歇性中辍后再次使用':
                continue
                
            series = df[col]
            q1 = series.quantile(0.25)
            median = series.median()
            q3 = series.quantile(0.75)
            iqr = q3 - q1
            
            # 计算校准锚点（使用Tukey's fences）
            fully_included = q3 + 1.5 * iqr
            cross_over = median
            fully_excluded = q1 - 1.5 * iqr
            
            # 存储校准参数
            self.results['calibration_anchors'][col] = {
                'fully_included': fully_included,
                'cross_over': cross_over,
                'fully_excluded': fully_excluded
            }
            
            # 执行校准
            calibrated_values = self._direct_calibration(series.values, 
                                                       fully_excluded, 
                                                       cross_over, 
                                                       fully_included)
            calibrated_df[col] = calibrated_values

        self.calibrated_data = calibrated_df
        print("模糊集校准完成")

    @staticmethod
    def _direct_calibration(x, excluded, crossover, included):
        """
        直接校准函数（S形函数）
        
        Args:
            x: 原始数值
            excluded: 完全不隶属锚点
            crossover: 交叉点
            included: 完全隶属锚点
            
        Returns:
            校准后的模糊集隶属度 (0-1之间)
        """
        result = np.zeros_like(x, dtype=float)
        
        for i, val in enumerate(x):
            if val <= excluded:
                result[i] = 0.0
            elif excluded < val <= crossover:
                result[i] = 0.5 * ((val - excluded) / (crossover - excluded))**2
            elif crossover < val < included:
                result[i] = 0.5 + 0.5 * ((val - crossover) / (included - crossover))**2
            else:  # val >= included
                result[i] = 1.0
                
        return result

    def necessity_analysis(self):
        """
        必要性分析
        计算每个条件变量对结果变量的一致性和覆盖度
        """
        if not self.calibrated_data:
            raise ValueError("请先执行calibrate_sets()")

        outcome_col = '用户间歇性中辍后再次使用'
        outcome = self.calibrated_data[outcome_col].values
        
        necessity_results = {}
        
        for col in self.calibrated_data.columns:
            if col == outcome_col:
                continue
                
            condition = self.calibrated_data[col].values
            
            # 计算一致性 consistency(X -> Y)
            numerator = np.minimum(condition, outcome).sum()
            denominator = condition.sum()
            consistency = numerator / denominator if denominator > 0 else 0
            
            # 计算覆盖度 coverage(X -> Y)
            cov_numerator = np.minimum(condition, outcome).sum()
            cov_denominator = outcome.sum()
            coverage = cov_numerator / cov_denominator if cov_denominator > 0 else 0
            
            necessity_results[col] = {
                'consistency': round(consistency, 3),
                'coverage': round(coverage, 3)
            }
        
        self.results['necessity_analysis'] = necessity_results
        print("必要性分析完成")

    def generate_truth_table(self, frequency_threshold=1, consistency_threshold=0.8):
        """
        生成真值表
        
        Args:
            frequency_threshold: 频数阈值
            consistency_threshold: 一致性阈值
        """
        if not self.calibrated_data:
            raise ValueError("请先执行calibrate_sets()")
            
        # 准备条件变量
        condition_cols = [col for col in self.calibrated_data.columns 
                         if col != '用户间歇性中辍后再次使用']
        conditions_matrix = self.calibrated_data[condition_cols].values
        outcome_vector = self.calibrated_data['用户间歇性中辍后再次使用'].values
        
        # 生成所有可能的组合（简化版，实际应用中应使用布尔运算）
        unique_configs = []
        config_id = 1
        
        for i in range(len(conditions_matrix)):
            # 简化处理：仅保留高于阈值的配置
            if outcome_vector[i] >= 0.5:  # 结果存在
                row = conditions_matrix[i]
                binary_row = (row >= 0.5).astype(int)  # 转换为二进制
                
                # 检查是否已存在相同配置
                is_duplicate = False
                for config in unique_configs:
                    if np.array_equal(config['binary'], binary_row):
                        config['frequency'] += 1
                        config['cases'].append(i)
                        is_duplicate = True
                        break
                        
                if not is_duplicate:
                    unique_configs.append({
                        'id': f'Config_{config_id}',
                        'binary': binary_row,
                        'original': row,
                        'frequency': 1,
                        'cases': [i]
                    })
                    config_id += 1
        
        # 过滤低频次配置
        filtered_configs = [c for c in unique_configs if c['frequency'] >= frequency_threshold]
        
        # 计算每种配置的一致性和覆盖度
        truth_table_data = []
        for config in filtered_configs:
            case_indices = config['cases']
            outcome_subset = outcome_vector[case_indices]
            condition_subset = config['original']  # 使用原始连续值计算一致性
            
            # 计算一致性
            min_vals = np.minimum(condition_subset, outcome_subset)
            consistency = min_vals.sum() / condition_subset.sum() if condition_subset.sum() > 0 else 0
            
            # 计算覆盖度
            coverage = min_vals.sum() / outcome_subset.sum() if outcome_subset.sum() > 0 else 0
            
            # 只保留高一致性配置
            if consistency >= consistency_threshold:
                row_data = {'组态ID': config['id']}
                for j, col in enumerate(condition_cols):
                    row_data[col] = config['binary'][j]
                row_data['一致性'] = round(consistency, 3)
                row_data['覆盖度'] = round(coverage, 3)
                row_data['频数'] = config['frequency']
                truth_table_data.append(row_data)
        
        self.results['truth_table'] = pd.DataFrame(truth_table_data)
        print(f"真值表生成完成，共识别出 {len(truth_table_data)} 个有效组态")

    def analyze_configurations(self):
        """
        组态分析
        识别核心条件与边缘条件
        """
        if self.results['truth_table'].empty:
            raise ValueError("请先执行generate_truth_table()")

        configs = self.results['truth_table']
        condition_cols = [col for col in configs.columns if col not in ['组态ID', '一致性', '覆盖度', '频数']]
        
        configuration_results = []
        
        for _, row in configs.iterrows():
            core_conditions = []
            peripheral_conditions = []
            
            # 简单判断：出现频率高的视为核心条件
            for col in condition_cols:
                if row[col] == 1:
                    # 在所有高一致性配置中的平均出现率
                    total_high_consistency = len(configs)
                    condition_support = configs[configs[col] == 1]
                    support_ratio = len(condition_support) / total_high_consistency
                    
                    if support_ratio >= 0.5:
                        core_conditions.append(col)
                    else:
                        peripheral_conditions.append(col)
            
            configuration_results.append({
                '组态': row['组态ID'],
                '一致性': row['一致性'],
                '覆盖度': row['覆盖度'],
                '核心条件': ', '.join(core_conditions),
                '边缘条件': ', '.join(peripheral_conditions),
                '解释': f"{row['组态ID']} 显示了导致再次使用的条件组合"
            })
        
        self.results['configurations'] = configuration_results
        print("组态分析完成")

    def export_results(self, output_path):
        """
        导出完整分析结果到Excel文件
        
        Args:
            output_path (str): 输出文件路径
        """
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 写入变量定义
            var_def_df = pd.DataFrame([
                {
                    '变量名称': name,
                    '变量含义': info.get('source_column', '复合变量'),
                    '数据来源': '问卷调查',
                    '测量方式': info.get('measurement', 'Likert量表均值')
                } for name, info in self.results['variable_definition'].items()
            ])
            var_def_df.to_excel(writer, sheet_name='变量定义', index=False)

            # 写入校准锚点
            calibration_df = pd.DataFrame([
                {
                    '变量名称': name,
                    '完全隶属': anchors['fully_included'],
                    '交叉点': anchors['cross_over'],
                    '完全不隶属': anchors['fully_excluded']
                } for name, anchors in self.results['calibration_anchors'].items()
            ])
            calibration_df.to_excel(writer, sheet_name='校准锚点', index=False)

            # 写入必要性分析
            necessity_df = pd.DataFrame([
                {
                    '条件变量': cond,
                    '一致性': metrics['consistency'],
                    '覆盖度': metrics['coverage']
                } for cond, metrics in self.results['necessity_analysis'].items()
            ]).sort_values('一致性', ascending=False)
            necessity_df.to_excel(writer, sheet_name='必要性分析', index=False)

            # 写入真值表
            self.results['truth_table'].to_excel(writer, sheet_name='真值表', index=False)

            # 写入组态分析
            config_df = pd.DataFrame(self.results['configurations'])
            config_df.to_excel(writer, sheet_name='组态分析', index=False)

            # 写入原始校准数据
            self.calibrated_data.to_excel(writer, sheet_name='校准后数据', index=False)

        print(f"完整分析结果已导出至: {output_path}")

    def run_complete_analysis(self, output_path):
        """
        执行完整的fsQCA分析流程
        
        Args:
            output_path (str): 结果输出路径
        """
        print("="*50)
        print("开始执行fsQCA完整分析流程")
        print("="*50)
        
        try:
            # 步骤1: 加载数据
            self.load_data()
            
            # 步骤2: 变量定义
            self.define_variables()
            
            # 步骤3: 模糊集校准
            self.calibrate_sets()
            
            # 步骤4: 必要性分析
            self.necessity_analysis()
            
            # 步骤5: 生成真值表
            self.generate_truth_table()
            
            # 步骤6: 组态分析
            self.analyze_configurations()
            
            # 步骤7: 导出结果
            self.export_results(output_path)
            
            print("="*50)
            print("fsQCA分析流程全部完成！")
            print("="*50)
            
        except Exception as e:
            print(f"分析过程中出现错误: {str(e)}")
            raise


def main():
    """
    主函数：演示完整的fsQCA分析流程
    """
    # 配置路径
    DATA_PATH = "/storage/data/7c0a973c-4c99-4242-914c-e0d8dca8d563/simulated_data.xlsx"
    OUTPUT_PATH = "/storage/data/7c0a973c-4c99-4242-914c-e0d8dca8d563/fsqca_analysis_complete.xlsx"
    
    # 创建分析器并执行完整流程
    analyzer = FsQCAAnalyzer(DATA_PATH)
    analyzer.run_complete_analysis(OUTPUT_PATH)


if __name__ == "__main__":
    main()


# 中文字体测试函数
def test_chinese_font_display():
    import matplotlib.pyplot as plt
    try:
        plt.figure(figsize=(8, 6))
        plt.bar(['认知负荷过高', '情感疏离感', '隐私侵犯担忧'], [0.8, 0.6, 0.7])
        plt.title('中文字体测试图', fontsize=16)
        plt.xlabel('条件变量', fontsize=14)
        plt.ylabel('隶属度', fontsize=14)
        plt.tight_layout()
        test_img_path = '/storage/data/eaa49533-e26a-4163-b3e5-afcd12f27092/chinese_font_test.png'
        plt.savefig(test_img_path, dpi=300, bbox_inches='tight')
        plt.close()
        print(f"中文字体测试图像已保存: {test_img_path}")
    except Exception as e:
        print(f"生成测试图像失败: {e}")

# 执行测试
test_chinese_font_display()

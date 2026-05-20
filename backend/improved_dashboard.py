"""
Improved Dashboard Generator for Azure Migration Assessment

This module creates an enhanced dashboard with:
1. Executive Summary
2. Move Status Breakdown by Resource Type
3. Top Blockers Analysis
4. Subscription Summary (if applicable)
5. Critical Actions Required

Usage:
    from backend.improved_dashboard import create_improved_dashboard
    create_improved_dashboard(file_path, resources, sub_map, friendly_map)
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import PieChart, Reference, Series, PieChart3D
from collections import Counter, defaultdict
import os
from datetime import datetime


def create_improved_dashboard(file_path, resources, sub_map, friendly_map):
    """
    Create an improved dashboard with comprehensive migration insights.
    
    Args:
        file_path: Path to the Excel file
        resources: List of resource dictionaries with migration analysis
        sub_map: Dictionary mapping subscription IDs to names
        friendly_map: Dictionary mapping resource types to friendly names
    """
    wb = load_workbook(file_path)
    
    # Remove existing Dashboard-Internal if present
    if "Dashboard-Internal" in wb.sheetnames:
        del wb["Dashboard-Internal"]
    
    # Remove existing Dashboard if present (old simple dashboard)
    if "Dashboard" in wb.sheetnames:
        del wb["Dashboard"]
    
    # Create new Dashboard sheet (improved version)
    ws = wb.create_sheet("Dashboard", 0)  # First sheet position
    
    # Define styles
    title_font = Font(size=16, bold=True, color='FFFFFF')
    title_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    
    header_font = Font(bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    
    subheader_font = Font(bold=True, size=12)
    subheader_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
    gray_fill = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')
    
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    center_align = Alignment(horizontal='center', vertical='center')
    
    # Analyze resources
    total_resources = len(resources)
    can_move = sum(1 for r in resources if r.get('Move Status') == 'Can Move')
    cannot_move = sum(1 for r in resources if r.get('Move Status') == 'Cannot Move')
    action_required = sum(1 for r in resources if r.get('Move Status') in ['Dependent - Action Required', 'Action Required'])
    review_required = sum(1 for r in resources if r.get('Move Status') == 'Review Required')
    
    # Calculate percentages
    can_move_pct = (can_move / total_resources * 100) if total_resources > 0 else 0
    cannot_move_pct = (cannot_move / total_resources * 100) if total_resources > 0 else 0
    action_pct = (action_required / total_resources * 100) if total_resources > 0 else 0
    review_pct = (review_required / total_resources * 100) if total_resources > 0 else 0
    
    current_row = 1
    
    # ==================== SECTION 1: TITLE ====================
    ws.merge_cells(f'A{current_row}:F{current_row}')
    title_cell = ws[f'A{current_row}']
    title_cell.value = "🚀 Azure Migration Assessment Dashboard"
    title_cell.font = title_font
    title_cell.fill = title_fill
    title_cell.alignment = center_align
    ws.row_dimensions[current_row].height = 30
    current_row += 2
    
    # ==================== SECTION 2: METADATA ====================
    ws[f'A{current_row}'] = "Assessment Date:"
    ws[f'B{current_row}'] = datetime.now().strftime("%Y-%m-%d %H:%M")
    ws[f'A{current_row}'].font = Font(bold=True)
    current_row += 1
    
    ws[f'A{current_row}'] = "Total Subscriptions:"
    ws[f'B{current_row}'] = len(sub_map)
    ws[f'A{current_row}'].font = Font(bold=True)
    current_row += 1
    
    ws[f'A{current_row}'] = "Assessment Type:"
    ws[f'B{current_row}'] = "Azure to Azure (Cross-Subscription)"
    ws[f'A{current_row}'].font = Font(bold=True)
    current_row += 1
    
    # Summary table headers
    headers = ['Metric', 'Count', 'Percentage']
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=current_row, column=col_idx)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center_align
    current_row += 1
    
    # Summary data
    summary_data = [
        ('Total Resources', total_resources, 1.0, None),
        ('✅ Can Move', can_move, can_move_pct / 100, green_fill),
        ('❌ Cannot Move', cannot_move, cannot_move_pct / 100, red_fill),
        ('⚠️ Can Move (Action Required)', action_required, action_pct / 100, yellow_fill),
        ('❓ Review Required', review_required, review_pct / 100, gray_fill)
    ]
    
    for label, count, pct, fill in summary_data:
        ws[f'A{current_row}'] = label
        ws[f'B{current_row}'] = count
        ws[f'C{current_row}'] = pct
        
        for col in ['A', 'B', 'C']:
            cell = ws[f'{col}{current_row}']
            cell.border = border
            if fill and col in ['B', 'C']:
                cell.fill = fill
            if col in ['B', 'C']:
                cell.alignment = center_align
            # Apply percentage format to column C
            if col == 'C':
                cell.number_format = '0.0%'
        current_row += 1
    
    current_row += 1
    

    
    # ==================== SECTION 4: MOVE STATUS BY RESOURCE TYPE ====================
    ws.merge_cells(f'A{current_row}:F{current_row}')
    type_summary_cell = ws[f'A{current_row}']
    type_summary_cell.value = "📦 Move Status Breakdown by Resource Type"
    type_summary_cell.font = subheader_font
    type_summary_cell.fill = subheader_fill
    type_summary_cell.border = border
    current_row += 1
    
    # Analyze by resource type
    type_analysis = defaultdict(lambda: {'total': 0, 'can_move': 0, 'cannot_move': 0, 'action': 0, 'review': 0})
    
    for res in resources:
        full_type = res.get("Resource Type", "").lower()
        friendly_type = friendly_map.get(full_type, full_type.split("/")[-1] if "/" in full_type else full_type)
        status = res.get('Move Status', 'Unknown')
        
        type_analysis[friendly_type]['total'] += 1
        if status == 'Can Move':
            type_analysis[friendly_type]['can_move'] += 1
        elif status == 'Cannot Move':
            type_analysis[friendly_type]['cannot_move'] += 1
        elif status in ['Dependent - Action Required', 'Action Required']:
            type_analysis[friendly_type]['action'] += 1
        elif status == 'Review Required':
            type_analysis[friendly_type]['review'] += 1
    
    # Headers
    type_headers = ['Resource Type', 'Total', 'Can Move', 'Cannot Move', 'Action Req.', 'Review Req.', '% Movable']
    for col_idx, header in enumerate(type_headers, start=1):
        cell = ws.cell(row=current_row, column=col_idx)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center_align
    current_row += 1
    
    # Data rows (sorted by total count descending)
    sorted_types = sorted(type_analysis.items(), key=lambda x: x[1]['total'], reverse=True)
    
    for res_type, stats in sorted_types:
        movable_pct = (stats['can_move'] / stats['total'] * 100) if stats['total'] > 0 else 0
        
        ws[f'A{current_row}'] = res_type
        ws[f'B{current_row}'] = stats['total']
        ws[f'C{current_row}'] = stats['can_move']
        ws[f'D{current_row}'] = stats['cannot_move']
        ws[f'E{current_row}'] = stats['action']
        ws[f'F{current_row}'] = stats['review']
        ws[f'G{current_row}'] = movable_pct / 100  # Write as decimal
        
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
            cell = ws[f'{col}{current_row}']
            cell.border = border
            if col in ['B', 'C', 'D', 'E', 'F', 'G']:
                cell.alignment = center_align
            
            # Color code and format the % Movable column
            if col == 'G':
                cell.number_format = '0.0%'  # Apply percentage format
                if movable_pct >= 90:
                    cell.fill = green_fill
                elif movable_pct >= 50:
                    cell.fill = yellow_fill
                else:
                    cell.fill = red_fill
        current_row += 1
    
    current_row += 1
    
    # ==================== SECTION 5: TOP BLOCKERS ====================
    ws.merge_cells(f'A{current_row}:F{current_row}')
    blockers_cell = ws[f'A{current_row}']
    blockers_cell.value = "🚫 Top Blocking Reasons"
    blockers_cell.font = subheader_font
    blockers_cell.fill = subheader_fill
    blockers_cell.border = border
    current_row += 1
    
    # Analyze blocking reasons
    blocker_counter = Counter()
    for res in resources:
        if res.get('Move Status') == 'Cannot Move':
            action = res.get('Action Required', 'Unknown reason')
            reason = action.split('.')[0][:100] if action else 'Unknown'
            blocker_counter[reason] += 1
    
    # Headers
    blocker_headers = ['Blocking Reason', 'Affected Resources']
    for col_idx, header in enumerate(blocker_headers, start=1):
        cell = ws.cell(row=current_row, column=col_idx)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center_align
    current_row += 1
    
    # Top 10 blockers
    top_blockers = blocker_counter.most_common(10)
    
    if top_blockers:
        for reason, count in top_blockers:
            ws[f'A{current_row}'] = reason
            ws[f'B{current_row}'] = count
            
            ws[f'A{current_row}'].border = border
            ws[f'B{current_row}'].border = border
            ws[f'B{current_row}'].alignment = center_align
            ws[f'B{current_row}'].fill = red_fill
            current_row += 1
    else:
        ws[f'A{current_row}'] = "No blocking reasons found"
        ws.merge_cells(f'A{current_row}:B{current_row}')
        ws[f'A{current_row}'].border = border
        ws[f'A{current_row}'].fill = green_fill
        current_row += 1
    
    current_row += 1
    
    # ==================== SECTION 6: CRITICAL ACTIONS REQUIRED ====================
    ws.merge_cells(f'A{current_row}:F{current_row}')
    action_cell = ws[f'A{current_row}']
    action_cell.value = "⚠️ Critical Actions Required"
    action_cell.font = subheader_font
    action_cell.fill = subheader_fill
    action_cell.border = border
    current_row += 1
    
    # Headers
    action_headers = ['Priority', 'Resource Name', 'Action', 'Resource Type']
    for col_idx, header in enumerate(action_headers, start=1):
        cell = ws.cell(row=current_row, column=col_idx)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center_align
    current_row += 1
    
    # Load non-movable resource types for grouping
    import json
    non_movable_types = set()
    try:
        non_movable_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data', 'non_movable_resources.json')
        # Fallback if that doesn't exist
        if not os.path.exists(non_movable_path):
            non_movable_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'non_movable_resources.json')
            
        with open(non_movable_path, 'r') as f:
            non_movable_data = json.load(f)
            
            # Handle new nested structure: {namespace: "Microsoft.X", types: ["typeA", "typeB"]}
            if 'non_movable_resources' in non_movable_data:
                for provider in non_movable_data['non_movable_resources']:
                    namespace = provider.get('namespace', '')
                    types = provider.get('types', [])
                    for resource_type in types:
                        # Build full type: "microsoft.insights/metricalerts"
                        full_type = f"{namespace}/{resource_type}".lower()
                        non_movable_types.add(full_type)
            
            # Also check for old format with 'non_movable_types' key
            elif isinstance(non_movable_data, dict):
                for provider_data in non_movable_data.values():
                    if isinstance(provider_data, dict) and 'non_movable_types' in provider_data:
                        non_movable_types.update(t.lower() for t in provider_data['non_movable_types'])
            
        print(f"📋 Loaded {len(non_movable_types)} non-movable resource types for grouping")
    except Exception as e:
        print(f"Warning: Could not load non_movable_resources.json: {e}")
        import traceback
        traceback.print_exc()
    
    # Collect actions - group non-movable, individual for others
    grouped_actions = {}  # For non-movable resources: {(priority, action, type): count}
    individual_actions = []  # For other resources: [(priority, name, action, type)]
    
    # 1. High Priority: Blocking Issues (Cannot Move)
    for res in resources:
        if res.get('Move Status') == 'Cannot Move':
            action = res.get('Action Required', 'Unknown blocking issue')
            res_name = res.get('Resource Name', 'Unknown')
            res_type = res.get('Resource Type', 'Unknown')
            res_type_lower = res_type.lower()
            
            # Check if this is a non-movable resource type (exact match)
            is_non_movable = res_type_lower in non_movable_types
            
            if is_non_movable:
                # Group non-movable resources
                # Extract friendly name from full type
                type_name = res_type.split('/')[-1] if '/' in res_type else res_type
                # Convert camelCase to Title Case (e.g., "metricAlerts" -> "Metric Alerts")
                import re
                friendly_type = re.sub(r'([a-z])([A-Z])', r'\1 \2', type_name).title()
                key = ('High', action, friendly_type)
                grouped_actions[key] = grouped_actions.get(key, 0) + 1
            else:
                # Show individually
                individual_actions.append(('High', res_name, action, res_type))
    
    # 2. High Priority: Critical Prep (Resources that need action before moving)
    for res in resources:
        if res.get('Move Status') == 'Action Required':
            action = res.get('Action Required', '')
            if action and any(k in action.lower() for k in ['detach', 'delete', 'remove', 'disable']):
                res_name = res.get('Resource Name', 'Unknown')
                res_type = res.get('Resource Type', 'Unknown')
                individual_actions.append(('High', res_name, action, res_type))
    
    # Note: Dependent - Action Required resources are NOT included here
    # because they don't need direct action - they're just waiting for other resources
    
    # Write grouped actions first
    if grouped_actions or individual_actions:
        # Write grouped non-movable resources
        for (priority, action, res_type), count in sorted(grouped_actions.items()):
            ws[f'A{current_row}'] = priority
            ws[f'B{current_row}'] = f"{res_type} ({count} resources)"
            ws[f'C{current_row}'] = action
            ws[f'D{current_row}'] = res_type
            
            # Styling
            for col in ['A', 'B', 'C', 'D']:
                cell = ws[f'{col}{current_row}']
                cell.border = border
                if col == 'A':
                    cell.alignment = center_align
            
            # Color coding for Priority
            if priority == 'High':
                ws[f'A{current_row}'].fill = red_fill
            elif priority == 'Medium':
                ws[f'A{current_row}'].fill = yellow_fill
                
            current_row += 1
        
        # Write individual actions
        for priority, res_name, action, res_type in individual_actions:
            ws[f'A{current_row}'] = priority
            ws[f'B{current_row}'] = res_name
            ws[f'C{current_row}'] = action
            ws[f'D{current_row}'] = res_type
            
            # Styling
            for col in ['A', 'B', 'C', 'D']:
                cell = ws[f'{col}{current_row}']
                cell.border = border
                if col == 'A':
                    cell.alignment = center_align
            
            # Color coding for Priority
            if priority == 'High':
                ws[f'A{current_row}'].fill = red_fill
            elif priority == 'Medium':
                ws[f'A{current_row}'].fill = yellow_fill
                
            current_row += 1
    else:
        ws[f'A{current_row}'] = "No critical actions detected"
        ws.merge_cells(f'A{current_row}:D{current_row}')
        ws[f'A{current_row}'].border = border
        ws[f'A{current_row}'].fill = green_fill
        current_row += 1
        
    current_row += 1
    
    # ==================== SECTION 7: SUBSCRIPTION SUMMARY ====================
    if len(sub_map) > 1:
        ws.merge_cells(f'A{current_row}:F{current_row}')
        sub_summary_cell = ws[f'A{current_row}']
        sub_summary_cell.value = "🔐 Subscription Summary"
        sub_summary_cell.font = subheader_font
        sub_summary_cell.fill = subheader_fill
        sub_summary_cell.border = border
        current_row += 1
        
        # Analyze by subscription
        sub_analysis = defaultdict(lambda: {'total': 0, 'can_move': 0, 'cannot_move': 0})
        
        for res in resources:
            sub_id = res.get('Subscription', 'Unknown')
            status = res.get('Move Status', 'Unknown')
            
            sub_analysis[sub_id]['total'] += 1
            if status == 'Can Move':
                sub_analysis[sub_id]['can_move'] += 1
            elif status == 'Cannot Move':
                sub_analysis[sub_id]['cannot_move'] += 1
        
        # Headers
        sub_headers = ['Subscription', 'Total Resources', 'Can Move', 'Cannot Move', '% Movable']
        for col_idx, header in enumerate(sub_headers, start=1):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.value = header
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
            cell.alignment = center_align
        current_row += 1
        
        # Data rows
        for sub_name, stats in sorted(sub_analysis.items()):
            movable_pct = (stats['can_move'] / stats['total'] * 100) if stats['total'] > 0 else 0
            
            ws[f'A{current_row}'] = sub_name
            ws[f'B{current_row}'] = stats['total']
            ws[f'C{current_row}'] = stats['can_move']
            ws[f'D{current_row}'] = stats['cannot_move']
            ws[f'E{current_row}'] = movable_pct / 100  # Write as decimal
            
            for col in ['A', 'B', 'C', 'D', 'E']:
                cell = ws[f'{col}{current_row}']
                cell.border = border
                if col in ['B', 'C', 'D', 'E']:
                    cell.alignment = center_align
                # Apply percentage format to column E
                if col == 'E':
                    cell.number_format = '0.0%'
            
            current_row += 1
    
    # ==================== AUTO-FIT COLUMNS ====================
    for col_idx in range(1, 7):
        col_letter = get_column_letter(col_idx)
        max_length = 0
        for row in ws[col_letter]:
            try:
                if row.value:
                    max_length = max(max_length, len(str(row.value)))
            except:
                pass
        adjusted_width = min(max_length + 2, 60)  # Cap at 60 for readability
        ws.column_dimensions[col_letter].width = adjusted_width
    
    # Save workbook
    wb.save(file_path)
    print(f"✅ Improved Dashboard created: {file_path}")


if __name__ == "__main__":
    print("This module should be imported and used within the main assessment script.")

import * as React from "react";
import { DefaultButton, PrimaryButton } from "@fluentui/react/lib/Button";
import {
	ColumnFilterGroup,
	ColumnFilterRule,
	FilterRuleMode,
	FilterInput,
	getFilterDefinitions,
	getFilterKind,
} from "./filtering";
import { fiscalInteger } from './dateConditions';

type FilterKind = "text" | "number" | "date" | "choice" | "boolean" | "lookup";

interface FilterDefinition {
	key: FilterRuleMode;
	text: string;
	input: FilterInput;
	types: string[];
}

interface FilterValueOption {
	label: string;
	value: string;
}

interface RuleState {
	mode: FilterRuleMode;
	value?: string | string[];
}

interface RuleValidation {
	message?: string;
	first?: string;
	second?: string;
}

export interface ColumnFilterEditorProps {
	columnName: string;
	dataType: string;
	initialGroup?: ColumnFilterGroup;
	valueOptions?: { label: string; value: string }[];
	onApply: (group: ColumnFilterGroup) => void;
	onCancel: () => void;
}

let nextEditorId = 0;

const DECIMAL_PATTERN = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/;
const DATE_PATTERN = /^(\d{4})-(\d{2})-(\d{2})$/;

function copyRuleValue(value: string | string[] | undefined): string | string[] | undefined {
	return Array.isArray(value) ? value.slice() : value;
}

function getRuleValues(value: string | string[] | undefined): [string, string] {
	if (Array.isArray(value)) return [value[0] ?? "", value[1] ?? ""];
	return [value ?? "", ""];
}

function getDefaultMode(dataType: string, definitions: FilterDefinition[]): FilterRuleMode {
	const kind = getFilterKind(dataType) as FilterKind;
	const preferredKeys = kind === "date" ? ["on", "equals"] : ["equals"];
	for (const key of preferredKeys) {
		const definition = definitions.find((candidate) => candidate.key === (key as FilterRuleMode));
		if (definition) return definition.key;
	}
	return definitions[0]?.key ?? ("equals" as FilterRuleMode);
}

function getDefinition(definitions: FilterDefinition[], mode: FilterRuleMode): FilterDefinition | undefined {
	return definitions.find((definition) => definition.key === mode);
}

function makeInitialRules(
	initialGroup: ColumnFilterGroup | undefined,
	defaultMode: FilterRuleMode
): RuleState[] {
	const initialRules = initialGroup?.rules ?? [];
	if (initialRules.length === 0) return [{ mode: defaultMode }];
	return initialRules.slice(0, 2).map((rule) => ({
		mode: rule.mode,
		value: copyRuleValue(rule.value),
	}));
}

function isValidDate(value: string): boolean {
	const match = DATE_PATTERN.exec(value);
	if (!match) return false;
	const year = Number(match[1]);
	const month = Number(match[2]);
	const day = Number(match[3]);
	const date = new Date(Date.UTC(year, month - 1, day));
	return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}

function isValidDecimal(value: string): boolean {
	return DECIMAL_PATTERN.test(value.trim()) && Number.isFinite(Number(value));
}

function isPositiveInteger(value: string): boolean {
	return /^\d+$/.test(value.trim()) && Number.isSafeInteger(Number(value)) && Number(value) > 0;
}

function isChoiceValue(value: string, options: FilterValueOption[]): boolean {
	return options.some((option) => option.value === value);
}

function validateScalar(
	value: string,
	definition: FilterDefinition,
	kind: FilterKind,
	valueOptions: FilterValueOption[]
): string | undefined {
	if (definition.input === 'period') return fiscalInteger(value, 13) ? undefined : 'Enter a fiscal period from 1 to 13.';
	if (definition.input === 'year') return fiscalInteger(value, 9999) ? undefined : 'Enter a fiscal year from 1 to 9999.';
	if (definition.input === "count") return isPositiveInteger(value) ? undefined : "Enter a positive whole number.";
	if (kind === "choice" || kind === "boolean" || kind === "lookup") {
		return isChoiceValue(value, valueOptions) ? undefined : "Select a value from the list.";
	}
	if (value.trim().length === 0) return "Enter a value.";
	if (kind === "number") return isValidDecimal(value) ? undefined : "Enter a valid number.";
	if (kind === "date") return isValidDate(value) ? undefined : "Enter a valid date.";
	return undefined;
}

function validateRule(
	rule: RuleState,
	definition: FilterDefinition | undefined,
	kind: FilterKind,
	valueOptions: FilterValueOption[]
): RuleValidation {
	if (!definition || definition.input === "none") return {};
	if (definition.input === 'periodYear') {
		const [period, year] = getRuleValues(rule.value);
		return {first: fiscalInteger(period, 13) ? undefined : 'Enter a fiscal period from 1 to 13.',
			second: fiscalInteger(year, 9999) ? undefined : 'Enter a fiscal year from 1 to 9999.'};
	}

	if (definition.input === "range") {
		const [first, second] = getRuleValues(rule.value);
		const firstError = validateScalar(first, { ...definition, input: "value" }, kind, valueOptions);
		const secondError = validateScalar(second, { ...definition, input: "value" }, kind, valueOptions);
		if (firstError || secondError) return { first: firstError, second: secondError };
		if (kind === "number" && Number(first) > Number(second)) {
			return { message: "The first value must be less than or equal to the second value." };
		}
		if (kind === "date" && first > second) {
			return { message: "The first date must be on or before the second date." };
		}
		return {};
	}

	const value = typeof rule.value === "string" ? rule.value : rule.value?.[0] ?? "";
	const error = validateScalar(value, definition, kind, valueOptions);
	return error ? { message: error } : {};
}

function errorId(editorId: string, ruleIndex: number, suffix: string): string {
	return `${editorId}-rule-${ruleIndex}-${suffix}`;
}

interface RuleEditorProps {
	editorId: string;
	ruleIndex: number;
	rule: RuleState;
	definition: FilterDefinition | undefined;
	definitions: FilterDefinition[];
	kind: FilterKind;
	valueOptions: FilterValueOption[];
	validation: RuleValidation;
	onModeChange: (mode: FilterRuleMode) => void;
	onValueChange: (value: string | string[] | undefined) => void;
}

const RuleEditor: React.FC<RuleEditorProps> = ({
	editorId,
	ruleIndex,
	rule,
	definition,
	definitions,
	kind,
	valueOptions,
	validation,
	onModeChange,
	onValueChange,
}) => {
	const modeId = errorId(editorId, ruleIndex, "mode");
	const valueId = errorId(editorId, ruleIndex, "value");
	const firstValueId = errorId(editorId, ruleIndex, "from");
	const secondValueId = errorId(editorId, ruleIndex, "to");
	const modeErrorId = errorId(editorId, ruleIndex, "mode-error");
	const valueErrorId = errorId(editorId, ruleIndex, "value-error");
	const firstErrorId = errorId(editorId, ruleIndex, "from-error");
	const secondErrorId = errorId(editorId, ruleIndex, "to-error");
	const rangeValues = getRuleValues(rule.value);
	const selectValue = typeof rule.value === "string" ? rule.value : rule.value?.[0] ?? "";

	const renderValueInput = (inputId: string, label: string, value: string, error?: string): React.ReactNode => {
		const integerInput = ['count', 'period', 'year', 'periodYear'].includes(definition?.input ?? '');
		const isDate = kind === "date" && !integerInput;
		const isNumber = kind === "number" || integerInput;
		const inputType = isDate ? "date" : isNumber ? "number" : "text";
		const step = integerInput ? "1" : isNumber ? "any" : undefined;
		return (
			<div className="cm-filter-editor-field">
				<label className="cm-filter-editor-label" htmlFor={inputId}>
					{label}
				</label>
				<input
					id={inputId}
					className="cm-filter-editor-input"
					type={inputType}
					value={value}
					min={integerInput ? 1 : undefined}
					step={step}
					inputMode={isNumber ? "decimal" : undefined}
					aria-invalid={error ? "true" : undefined}
					aria-describedby={error ? (inputId === firstValueId ? firstErrorId : inputId === secondValueId ? secondErrorId : valueErrorId) : undefined}
					onChange={(event: React.ChangeEvent<HTMLInputElement>) => {
						const nextValue = event.target.value;
						onValueChange(inputId === firstValueId ? [nextValue, rangeValues[1]] : inputId === secondValueId ? [rangeValues[0], nextValue] : nextValue);
					}}
				/>
				{error && (
					<div id={inputId === firstValueId ? firstErrorId : inputId === secondValueId ? secondErrorId : valueErrorId} className="cm-filter-editor-error" role="alert">
						{error}
					</div>
				)}
			</div>
		);
	};

	return (
		<div className="cm-filter-editor-rule">
			<div className="cm-filter-editor-field">
				<label className="cm-filter-editor-label" htmlFor={modeId}>
					Condition
				</label>
				<select
					id={modeId}
					className="cm-filter-editor-select"
					value={rule.mode as string}
					aria-invalid={validation.message ? "true" : undefined}
					aria-describedby={validation.message ? modeErrorId : undefined}
					onChange={(event: React.ChangeEvent<HTMLSelectElement>) => onModeChange(event.target.value as FilterRuleMode)}
				>
					{definitions.map((candidate) => (
						<option key={candidate.key as string} value={candidate.key as string}>
							{candidate.text}
						</option>
					))}
				</select>
				{validation.message && definition?.input === "none" && (
					<div id={modeErrorId} className="cm-filter-editor-error" role="alert">
						{validation.message}
					</div>
				)}
			</div>

			{definition?.input === 'periodYear' ? (
				<div className="cm-filter-editor-range">
					{renderValueInput(firstValueId, 'Fiscal period', rangeValues[0], validation.first)}
					{renderValueInput(secondValueId, 'Fiscal year', rangeValues[1], validation.second)}
				</div>
			) : definition?.input === "range" && (kind === "choice" || kind === "boolean" || kind === "lookup") ? (
				<div className="cm-filter-editor-range">
					<div className="cm-filter-editor-field">
						<label className="cm-filter-editor-label" htmlFor={firstValueId}>
							From
						</label>
						<select
							id={firstValueId}
							className="cm-filter-editor-select"
							value={rangeValues[0]}
							aria-invalid={validation.first ? "true" : undefined}
							onChange={(event: React.ChangeEvent<HTMLSelectElement>) => onValueChange([event.target.value, rangeValues[1]])}
						>
							<option value="">Select a value</option>
							{valueOptions.map((option) => (
								<option key={option.value} value={option.value}>
									{option.label}
								</option>
							))}
						</select>
						{validation.first && <div className="cm-filter-editor-error" role="alert">{validation.first}</div>}
					</div>
					<div className="cm-filter-editor-field">
						<label className="cm-filter-editor-label" htmlFor={secondValueId}>
							To
						</label>
						<select
							id={secondValueId}
							className="cm-filter-editor-select"
							value={rangeValues[1]}
							aria-invalid={validation.second ? "true" : undefined}
							onChange={(event: React.ChangeEvent<HTMLSelectElement>) => onValueChange([rangeValues[0], event.target.value])}
						>
							<option value="">Select a value</option>
							{valueOptions.map((option) => (
								<option key={option.value} value={option.value}>
									{option.label}
								</option>
							))}
						</select>
						{validation.second && <div className="cm-filter-editor-error" role="alert">{validation.second}</div>}
					</div>
				</div>
			) : definition?.input === "range" ? (
				<div className="cm-filter-editor-range">
					{renderValueInput(firstValueId, kind === "date" ? "From" : "Minimum", rangeValues[0], validation.first)}
					{renderValueInput(secondValueId, kind === "date" ? "To" : "Maximum", rangeValues[1], validation.second)}
					{validation.message && <div className="cm-filter-editor-error" role="alert">{validation.message}</div>}
				</div>
			) : definition?.input !== "none" && (kind === "choice" || kind === "boolean" || kind === "lookup") ? (
				<div className="cm-filter-editor-field">
					<label className="cm-filter-editor-label" htmlFor={valueId}>
						Value
					</label>
					<select
						id={valueId}
						className="cm-filter-editor-select"
						value={selectValue}
						aria-invalid={validation.message ? "true" : undefined}
						aria-describedby={validation.message ? valueErrorId : undefined}
						onChange={(event: React.ChangeEvent<HTMLSelectElement>) => onValueChange(event.target.value)}
					>
						<option value="">Select a value</option>
						{valueOptions.map((option) => (
							<option key={option.value} value={option.value}>
								{option.label}
							</option>
						))}
					</select>
					{validation.message && <div id={valueErrorId} className="cm-filter-editor-error" role="alert">{validation.message}</div>}
				</div>
			) : definition && ['count', 'period', 'year'].includes(definition.input) ? (
				renderValueInput(valueId, definition.input === 'period' ? 'Fiscal period' : definition.input === 'year' ? 'Fiscal year' : 'Count', selectValue, validation.message)
			) : definition?.input === "value" ? (
				renderValueInput(valueId, "Value", selectValue, validation.message)
			) : null}
		</div>
	);
};

function makeRuleForApply(rule: RuleState, definition: FilterDefinition | undefined): ColumnFilterRule {
	const result: ColumnFilterRule = { mode: rule.mode };
	if (definition?.input !== "none" && rule.value !== undefined) result.value = copyRuleValue(rule.value);
	return result;
}

export const ColumnFilterEditor: React.FC<ColumnFilterEditorProps> = ({
	columnName,
	dataType,
	initialGroup,
	valueOptions = [],
	onApply,
	onCancel,
}) => {
	const definitions = getFilterDefinitions(dataType) as FilterDefinition[];
	const kind = getFilterKind(dataType) as FilterKind;
	const defaultMode = getDefaultMode(dataType, definitions);
	const editorIdRef = React.useRef<string | undefined>(undefined);
	if (!editorIdRef.current) {
		nextEditorId += 1;
		editorIdRef.current = `cm-filter-editor-${nextEditorId}`;
	}
	const editorId = editorIdRef.current;
	const [rules, setRules] = React.useState<RuleState[]>(() => makeInitialRules(initialGroup, defaultMode));
	const [operator, setOperator] = React.useState<"and" | "or">(() => initialGroup?.operator === "or" ? "or" : "and");
	const [hasSecondCondition, setHasSecondCondition] = React.useState<boolean>(() => (initialGroup?.rules.length ?? 0) > 1);
	const [validation, setValidation] = React.useState<Record<number, RuleValidation>>({});
	const [applyError, setApplyError] = React.useState<string | undefined>(undefined);

	const clearErrors = (): void => {
		setValidation({});
		setApplyError(undefined);
	};

	const updateRule = (ruleIndex: number, update: Partial<RuleState>): void => {
		setRules((current) => current.map((rule, index) => index === ruleIndex ? { ...rule, ...update } : rule));
		clearErrors();
	};

	const addSecondRule = (): void => {
		setHasSecondCondition((current) => {
			if (!current) {
				setRules((existing) => existing.length > 1 ? existing : [...existing, { mode: defaultMode }]);
			}
			return !current;
		});
		clearErrors();
	};

	const handleApply = (): void => {
		const activeRules = hasSecondCondition ? rules.slice(0, 2) : rules.slice(0, 1);
		const nextValidation: Record<number, RuleValidation> = {};
		activeRules.forEach((rule, index) => {
			const result = validateRule(rule, getDefinition(definitions, rule.mode), kind, valueOptions);
			if (result.message || result.first || result.second) nextValidation[index] = result;
		});
		if (Object.keys(nextValidation).length > 0) {
			setValidation(nextValidation);
			setApplyError(undefined);
			return;
		}

		const group: ColumnFilterGroup = {
			operator,
			rules: activeRules.map((rule) => makeRuleForApply(rule, getDefinition(definitions, rule.mode))),
		};
		try {
			onApply(group);
			setValidation({});
			setApplyError(undefined);
		} catch (error) {
			setApplyError(error instanceof Error ? error.message : String(error));
		}
	};

	return (
		<div className="cm-filter-editor">
			<div className="cm-filter-editor-header">
				<h2 className="cm-filter-editor-title" style={{ margin: 0 }}>
					{columnName} filters
				</h2>
				<DefaultButton className="cm-filter-editor-back" text="Back" onClick={onCancel} />
			</div>

			<div className="cm-filter-editor-body">
			<div className="cm-filter-note cm-filter-editor-note">
				Conditions on other columns are combined with AND.
			</div>

			<RuleEditor
				editorId={editorId}
				ruleIndex={0}
				rule={rules[0]}
				definition={getDefinition(definitions, rules[0].mode)}
				definitions={definitions}
				kind={kind}
				valueOptions={valueOptions}
				validation={validation[0] ?? {}}
				onModeChange={(mode) => updateRule(0, { mode, value: undefined })}
				onValueChange={(value) => updateRule(0, { value })}
			/>

			<div className="cm-filter-editor-second-toggle">
				<label className="cm-filter-editor-checkbox-label">
					<input
						type="checkbox"
						checked={hasSecondCondition}
						onChange={addSecondRule}
					/>
					<span>Add a second condition</span>
				</label>
			</div>

			{hasSecondCondition && (
				<div className="cm-filter-editor-second-condition">
					<div className="cm-filter-editor-field">
						<label className="cm-filter-editor-label" htmlFor={`${editorId}-operator`}>
							Combine conditions
						</label>
						<select
							id={`${editorId}-operator`}
							className="cm-filter-editor-select"
							value={operator}
							onChange={(event: React.ChangeEvent<HTMLSelectElement>) => {
								setOperator(event.target.value === "or" ? "or" : "and");
								clearErrors();
							}}
						>
							<option value="and">AND</option>
							<option value="or">OR</option>
						</select>
					</div>
					<RuleEditor
						editorId={editorId}
						ruleIndex={1}
						rule={rules[1] ?? { mode: defaultMode }}
						definition={getDefinition(definitions, (rules[1] ?? { mode: defaultMode }).mode)}
						definitions={definitions}
						kind={kind}
						valueOptions={valueOptions}
						validation={validation[1] ?? {}}
						onModeChange={(mode) => updateRule(1, { mode, value: undefined })}
						onValueChange={(value) => updateRule(1, { value })}
					/>
				</div>
			)}

			{applyError && <div className="cm-filter-editor-error cm-filter-editor-apply-error" role="alert">{applyError}</div>}
			</div>

			<div className="cm-filter-actions cm-filter-editor-actions">
				<PrimaryButton text="Apply" onClick={handleApply} />
				<DefaultButton text="Cancel" onClick={onCancel} />
			</div>
		</div>
	);
};

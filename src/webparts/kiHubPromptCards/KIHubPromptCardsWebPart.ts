import { Version } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PropertyPaneTextField, PropertyPaneDropdown, IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration
} from '@microsoft/sp-webpart-base';

import styles from './KIHubPromptCardsWebPart.module.scss';

export interface IKIHubPromptCardsWebPartProps {
  listName: string;
  copilotUrl: string;
  filterType: string;
  filterValue: string;
}

interface IPromptLinkValue {
  Url?: string;
  Description?: string;
}

interface IPromptItem {
  Id: number;
  Title: string;
  field_1?: string;
  field_2?: string;
  ProgramAreas?: string;
  Featured?: boolean;
  PromptLink?: string | IPromptLinkValue;
}

export default class KIHubPromptCardsWebPart extends BaseClientSideWebPart<IKIHubPromptCardsWebPartProps> {
  public async render(): Promise<void> {
    this.domElement.innerHTML = `
      <section class="${styles.promptCards}">
        <div class="${styles.headerBlock}">
          <div class="${styles.kicker}">Copilot Prompt Library</div>
          <h2 class="${styles.mainTitle}">Prompt Cards</h2>
          <p class="${styles.mainSubtitle}">
            Start with a polished prompt, copy it quickly, and launch Copilot in one click.
          </p>
        </div>

        <div class="${styles.loadingState}">
          Loading prompts...
        </div>
      </section>
    `;

    try {
      const items: IPromptItem[] = await this._getPromptItems();
      this._renderCards(items);
      this._wireUpEvents();
    } catch (error) {
      console.error(error);
      this.domElement.innerHTML = `
        <section class="${styles.promptCards}">
          <div class="${styles.errorState}">
            Unable to load the Copilot Prompt Library right now.
          </div>
        </section>
      `;
    }
  }

  private async _getPromptItems(): Promise<IPromptItem[]> {
    const listName: string = this.properties.listName || 'Copilot Prompt Library';
    const filterType: string = this.properties.filterType || 'None';
    const filterValue: string = this.properties.filterValue || '';

    let filterQuery: string = '';

    if (filterType === 'Featured') {
      filterQuery = `&$filter=Featured eq 1`;
    } else if (filterType === 'ProgramArea' && filterValue) {
      filterQuery = `&$filter=ProgramAreas eq '${filterValue.replace(/'/g, "''")}'`;
    }

    const endpoint: string =
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items` +
      `?$select=Id,Title,field_1,field_2,ProgramAreas,Featured,PromptLink` +
      `${filterQuery}` +
      `&$orderby=Featured desc,Id asc`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Error loading prompt items: ${response.status} ${response.statusText}`);
    }

    const json: { value: IPromptItem[] } = await response.json();
    return json.value || [];
  }

  private _escapeHtml(value?: string): string {
    if (!value) {
      return '';
    }

    return value
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  private _truncate(value: string, maxLength: number): string {
    if (!value) {
      return '';
    }

    return value.length > maxLength ? `${value.substring(0, maxLength)}...` : value;
  }

  private _resolvePromptLink(item: IPromptItem): string {
    const fallback: string = this.properties.copilotUrl || 'https://m365.cloud.microsoft/chat';
    const raw: string | IPromptLinkValue | undefined = item.PromptLink;

    if (!raw) {
      return fallback;
    }

    if (typeof raw === 'string') {
      return raw.startsWith('http') ? raw : fallback;
    }

    if (raw.Url && raw.Url.startsWith('http')) {
      return raw.Url;
    }

    return fallback;
  }

  private _renderCards(items: IPromptItem[]): void {
    if (!items.length) {
      this.domElement.innerHTML = `
        <section class="${styles.promptCards}">
          <div class="${styles.emptyState}">
            No prompt items matched this page filter.
          </div>
        </section>
      `;
      return;
    }

    const cardsHtml: string = items.map((item: IPromptItem) => {
      const title: string = this._escapeHtml(item.Title || 'Untitled Prompt');
      const category: string = this._escapeHtml(item.ProgramAreas || 'General');
      const beginnerPrompt: string = this._escapeHtml(item.field_1 || '');
      const advancedPrompt: string = this._escapeHtml(item.field_2 || '');
      const featured: boolean = !!item.Featured;
      const promptLink: string = this._escapeHtml(this._resolvePromptLink(item));

      const initialPrompt: string = beginnerPrompt || advancedPrompt || 'No prompt text available.';
      const previewPrompt: string = this._truncate(initialPrompt, 320);

      return `
        <article class="${styles.card}">
          <div class="${styles.cardHeader}">
            <div class="${styles.badgeRow}">
              <span class="${styles.categoryBadge}">${category}</span>
              ${featured ? `<span class="${styles.featuredBadge}">Featured</span>` : ''}
            </div>

            <h3 class="${styles.cardTitle}">${title}</h3>
          </div>

          <div class="${styles.modeRow}">
            <button
              type="button"
              class="${styles.modeButton} ${styles.modeButtonActive}"
              data-role="mode"
              data-mode="beginner"
              data-card-id="${item.Id}">
              Beginner
            </button>

            <button
              type="button"
              class="${styles.modeButton}"
              data-role="mode"
              data-mode="advanced"
              data-card-id="${item.Id}">
              Advanced
            </button>
          </div>

          <div
            class="${styles.promptBody}"
            id="prompt-body-${item.Id}"
            data-beginner="${beginnerPrompt}"
            data-advanced="${advancedPrompt}"
            data-current="${initialPrompt}">
            ${previewPrompt}
          </div>

          <div class="${styles.actionRow}">
            <button
              type="button"
              class="${styles.secondaryButton}"
              data-role="copy"
              data-card-id="${item.Id}">
              Copy Prompt
            </button>

            <button
              type="button"
              class="${styles.primaryButton}"
              data-role="copilot"
              data-card-id="${item.Id}"
              data-link="${promptLink}">
              Use in Copilot
            </button>
          </div>
        </article>
      `;
    }).join('');

    this.domElement.innerHTML = `
      <section class="${styles.promptCards}">
        <div class="${styles.headerBlock}">
          <div class="${styles.kicker}">Copilot Prompt Library</div>
          <h2 class="${styles.mainTitle}">Prompt Cards</h2>
          <p class="${styles.mainSubtitle}">
            Start with a polished prompt, copy it quickly, and launch Copilot in one click.
          </p>
        </div>

        <div class="${styles.grid}">
          ${cardsHtml}
        </div>

        <div id="kihub-toast" class="${styles.toast}" aria-live="polite"></div>
      </section>
    `;
  }

  private _wireUpEvents(): void {
    const modeButtons: NodeListOf<HTMLButtonElement> =
      this.domElement.querySelectorAll('button[data-role="mode"]');

    modeButtons.forEach((button: HTMLButtonElement) => {
      button.addEventListener('click', () => {
        const cardId: string = button.getAttribute('data-card-id') || '';
        const mode: string = button.getAttribute('data-mode') || 'beginner';
        const promptBody: HTMLElement | null = this.domElement.querySelector(`#prompt-body-${cardId}`);

        if (!promptBody) {
          return;
        }

        const beginnerPrompt: string = promptBody.getAttribute('data-beginner') || '';
        const advancedPrompt: string = promptBody.getAttribute('data-advanced') || '';

        let selectedPrompt: string = '';
        if (mode === 'advanced') {
          selectedPrompt = advancedPrompt || beginnerPrompt || 'No prompt text available.';
        } else {
          selectedPrompt = beginnerPrompt || advancedPrompt || 'No prompt text available.';
        }

        promptBody.setAttribute('data-current', selectedPrompt);
        promptBody.textContent = this._truncate(selectedPrompt, 320);

        const siblingButtons: NodeListOf<HTMLButtonElement> =
          this.domElement.querySelectorAll(`button[data-role="mode"][data-card-id="${cardId}"]`);

        siblingButtons.forEach((sibling: HTMLButtonElement) => {
          sibling.classList.remove(styles.modeButtonActive);
        });

        button.classList.add(styles.modeButtonActive);
      });
    });

    const copyButtons: NodeListOf<HTMLButtonElement> =
      this.domElement.querySelectorAll('button[data-role="copy"]');

    copyButtons.forEach((button: HTMLButtonElement) => {
      button.addEventListener('click', async () => {
        const cardId: string = button.getAttribute('data-card-id') || '';
        const promptBody: HTMLElement | null = this.domElement.querySelector(`#prompt-body-${cardId}`);
        const prompt: string = promptBody?.getAttribute('data-current') || '';

        await this._copyPrompt(prompt);
      });
    });

    const copilotButtons: NodeListOf<HTMLButtonElement> =
      this.domElement.querySelectorAll('button[data-role="copilot"]');

    copilotButtons.forEach((button: HTMLButtonElement) => {
      button.addEventListener('click', async () => {
        const cardId: string = button.getAttribute('data-card-id') || '';
        const promptBody: HTMLElement | null = this.domElement.querySelector(`#prompt-body-${cardId}`);
        const prompt: string = promptBody?.getAttribute('data-current') || '';
        const targetUrl: string =
          button.getAttribute('data-link') ||
          this.properties.copilotUrl ||
          'https://m365.cloud.microsoft/chat';

        await this._copyPrompt(prompt);
        this._showToast('Prompt copied. Opening Copilot...');
        window.open(targetUrl, '_blank');
      });
    });
  }

  private async _copyPrompt(prompt: string): Promise<void> {
    try {
      await navigator.clipboard.writeText(prompt);
      this._showToast('Prompt copied.');
    } catch (error) {
      console.error('Clipboard copy failed.', error);
      this._showToast('Copy failed. Please copy manually.');
    }
  }

  private _showToast(message: string): void {
    const toast: HTMLElement | null = this.domElement.querySelector('#kihub-toast');

    if (!toast) {
      return;
    }

    toast.textContent = message;
    toast.classList.add(styles.toastVisible);

    window.setTimeout(() => {
      toast.classList.remove(styles.toastVisible);
    }, 2200);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const filterOptions: IPropertyPaneDropdownOption[] = [
      { key: 'None', text: 'None' },
      { key: 'Featured', text: 'Featured only' },
      { key: 'ProgramArea', text: 'Program Area' }
    ];

    return {
      pages: [
        {
          header: {
            description: 'Prompt Cards Settings'
          },
          groups: [
            {
              groupName: 'Data',
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: 'List Name'
                }),
                PropertyPaneTextField('copilotUrl', {
                  label: 'Default Copilot URL'
                }),
                PropertyPaneDropdown('filterType', {
                  label: 'Filter Type',
                  options: filterOptions
                }),
                PropertyPaneTextField('filterValue', {
                  label: 'Filter Value'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
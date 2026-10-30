
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

export interface IEmployeeRequestsBannerApplicationCustomizerProperties {
  testMessage: string;
  BannerMessage: string;
}

export default class EmployeeRequestsBannerApplicationCustomizer
  extends BaseApplicationCustomizer<IEmployeeRequestsBannerApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;

  private readonly _bannerClassName = 'employee-requests-banner';

  public onInit(): Promise<void> {

    // Re-render when SharePoint placeholders change.
    this.context.placeholderProvider.changedEvent.add(
      this,
      this._renderPlaceholders
    );

    // Re-render after client-side navigation.
    this.context.application.navigatedEvent.add(
      this,
      this._renderPlaceholders
    );

    // Initial render.
    this._renderPlaceholders();

    return Promise.resolve();
  }

  /**
   * Determines whether the current page is the Employee list.
   * Update this condition to match your actual list URL.
   */
  private _isEmployeeListPage(): boolean {
    const currentUrl = window.location.pathname.toLowerCase();

    return currentUrl.includes('/lists/employee/');
  }

  /**
   * Creates or updates the banner in the Top placeholder.
   */
  private _renderPlaceholders = (): void => {
    // Ignore a cached placeholder whose DOM element was detached.
    if (
      this._topPlaceholder &&
      !this._topPlaceholder.domElement.isConnected
    ) {
      this._topPlaceholder = undefined;
    }

    // Reuse the existing placeholder or try to create it.
    if (!this._topPlaceholder) {
      this._topPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose }
        );
    }

    // The Top placeholder may be unavailable on some pages.
    if (!this._topPlaceholder) {
      return;
    }

    const container = this._topPlaceholder.domElement;

    // Find the existing banner.
    let banner = container.querySelector(
      `.${this._bannerClassName}`
    ) as HTMLDivElement | null;

    // Create the banner only once per placeholder.
    if (!banner) {
      banner = document.createElement('div');

      banner.className = this._bannerClassName;

      banner.textContent =
        this.properties.BannerMessage ||
        'Welcome to our SharePoint Portal!';

      banner.style.backgroundColor = '#0078d4';
      banner.style.color = '#ffffff';
      banner.style.textAlign = 'center';
      banner.style.padding = '10px';
      banner.style.fontSize = '14px';
      banner.style.fontWeight = 'bold';

      container.appendChild(banner);
    }

    // Hide only our banner on the Employee list page.
    const shouldHideBanner = this._isEmployeeListPage();

    banner.style.display = shouldHideBanner ? 'none' : '';
  };

  /**
   * Called when SharePoint disposes of the placeholder.
   */
  private _onDispose = (): void => {

    // Clear the reference so it can be recreated.
    this._topPlaceholder = undefined;
  };
}

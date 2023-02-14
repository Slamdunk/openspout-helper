CSFIX_PHP_BIN=PHP_CS_FIXER_IGNORE_ENV=1 php8.2
PHP_BIN=php8.2 -d zend.assertions=1
COMPOSER_BIN=$(shell command -v composer)

SRCS := $(shell find ./lib ./tests -type f -not -path "*/tmp/*")

LOCAL_BASE_BRANCH ?= $(shell git show-branch | sed "s/].*//" | grep "\*" | grep -v "$$(git rev-parse --abbrev-ref HEAD)" | head -n1 | sed "s/^.*\[//")
ifeq ($(strip $(LOCAL_BASE_BRANCH)),)
	LOCAL_BASE_BRANCH := HEAD^
endif
BASE_BRANCH ?= $(LOCAL_BASE_BRANCH)

all: csfix static-analysis code-coverage
	@echo "Done."

vendor: composer.json
	$(PHP_BIN) $(COMPOSER_BIN) update
	$(PHP_BIN) $(COMPOSER_BIN) bump
	touch vendor

.PHONY: csfix
csfix: vendor
	$(CSFIX_PHP_BIN) vendor/bin/php-cs-fixer fix -v $(arg)

.PHONY: static-analysis
static-analysis: vendor
	$(PHP_BIN) vendor/bin/phpstan analyse $(PHPSTAN_ARGS)

coverage/junit.xml: vendor $(SRCS) Makefile
	$(PHP_BIN) vendor/bin/phpunit $(PHPUNIT_ARGS)

.PHONY: test
test: coverage/junit.xml

.PHONY: code-coverage
code-coverage: coverage/junit.xml
	echo "Base branch: $(BASE_BRANCH)"
	$(PHP_BIN) \
		vendor/bin/infection \
		--threads=$(shell nproc) \
		--git-diff-lines \
		--git-diff-base=$(BASE_BRANCH) \
		--skip-initial-tests \
		--coverage=coverage \
		--show-mutations \
		--verbose \
		--min-msi=100 \
		$(INFECTION_ARGS)
